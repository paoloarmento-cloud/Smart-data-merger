"""
Excel/CSV Merger Tool - Core Module
Handles file operations, key detection, and merge logic
"""

import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
from typing import List, Tuple, Dict, Optional
import os
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class MergeEngine:
    """Core engine for Excel/CSV file merging operations"""
    
    def __init__(self):
        self.df1 = None
        self.df2 = None
        self.file1_path = ""
        self.file2_path = ""
        self.merge_result = None
        self.detected_keys = []
        
    def load_file(self, file_path: str, file_number: int) -> bool:
        """
        Load Excel or CSV file into DataFrame
        Args:
            file_path: Path to the file
            file_number: 1 or 2 to indicate which file slot
        Returns:
            Boolean indicating success
        """
        try:
            file_ext = os.path.splitext(file_path)[1].lower()
            
            if file_ext in ['.xlsx', '.xls']:
                df = pd.read_excel(file_path, sheet_name=0)  # First sheet only
            elif file_ext == '.csv':
                # Try different encodings
                for encoding in ['utf-8', 'utf-16', 'windows-1252']:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    logger.error(f"Could not decode CSV file: {file_path}")
                    return False
            elif file_ext == '.txt':
                # Assume tab-delimited for .txt files
                df = pd.read_csv(file_path, sep='\t', encoding='utf-8')
            else:
                logger.error(f"Unsupported file format: {file_ext}")
                return False
            
            # Basic validation
            if df.empty:
                logger.error(f"File is empty: {file_path}")
                return False
            
            # Clean column names (strip whitespace)
            df.columns = df.columns.str.strip()
            
            # Store the DataFrame
            if file_number == 1:
                self.df1 = df
                self.file1_path = file_path
            else:
                self.df2 = df
                self.file2_path = file_path
                
            logger.info(f"Loaded file {file_number}: {os.path.basename(file_path)} "
                       f"({len(df)} rows, {len(df.columns)} columns)")
            return True
            
        except Exception as e:
            logger.error(f"Error loading file {file_path}: {str(e)}")
            return False
    
    def detect_merge_keys(self, min_match_ratio: float = 0.3) -> List[Tuple[str, str, float]]:
        """
        Automatically detect potential merge keys between two DataFrames
        Args:
            min_match_ratio: Minimum ratio of matching values required
        Returns:
            List of tuples (col1, col2, match_ratio) sorted by match quality
        """
        if self.df1 is None or self.df2 is None:
            logger.error("Both files must be loaded before detecting keys")
            return []
        
        candidates = []
        # Get columns with reasonable uniqueness in df1
        df1_unique_cols = []
        for col in self.df1.columns:
            unique_ratio = self.df1[col].nunique() / len(self.df1)
            if unique_ratio > 0.7:  # At least 70% unique values
                df1_unique_cols.append(col)
        
        # Get columns with reasonable uniqueness in df2  
        df2_unique_cols = []
        for col in self.df2.columns:
            unique_ratio = self.df2[col].nunique() / len(self.df2)
            if unique_ratio > 0.7:  # At least 70% unique values
                df2_unique_cols.append(col)
        
        # Check combinations of unique columns
        for col1 in df1_unique_cols:
            for col2 in df2_unique_cols:
                # Calculate column name similarity
                name_similarity = fuzz.ratio(col1.lower(), col2.lower()) / 100
                
                # Calculate value overlap
                # Calculate value overlap
                try:
                    
                    # Convert to string and clean for comparison (same normalization as merge)
                    values1 = set(self.df1[col1].astype(str).str.strip().str.replace('.0', '').str.upper())
                    values2 = set(self.df2[col2].astype(str).str.strip().str.replace('.0', '').str.upper())
    
                    # Remove empty/null values
                    values1 = {v for v in values1 if v not in ['', 'NAN', 'NONE', 'NULL']}
                    values2 = {v for v in values2 if v not in ['', 'NAN', 'NONE', 'NULL']}
    
                    intersection = values1.intersection(values2)
                    overlap_ratio = len(intersection) / min(len(values1), len(values2))
                    combined_score = (name_similarity * 0.3) + (overlap_ratio * 0.7)
    
                    # Convert to string and clean for comparison
                    values1 = set(self.df1[col1].astype(str).str.strip().str.replace('.0', '').str.upper())
                    values2 = set(self.df2[col2].astype(str).str.strip().str.replace('.0', '').str.upper())
                    
                    # Remove empty/null values
                    values1 = {v for v in values1 if v not in ['', 'NAN', 'NONE', 'NULL']}
                    values2 = {v for v in values2 if v not in ['', 'NAN', 'NONE', 'NULL']}
                    
                    if len(values1) == 0 or len(values2) == 0:
                        continue
                        
                    # Calculate overlap ratio
                    intersection = values1.intersection(values2)
                    overlap_ratio = len(intersection) / min(len(values1), len(values2))
                    
                    # Combined score: weighted average of name similarity and value overlap
                    combined_score = (name_similarity * 0.3) + (overlap_ratio * 0.7)
                    
                    if combined_score >= min_match_ratio:
                        candidates.append((col1, col2, combined_score, len(intersection)))
                        
                except Exception as e:
                    logger.warning(f"Error comparing columns {col1} and {col2}: {str(e)}")
                    continue
        
        # Sort by combined score (descending) and number of matches
        candidates.sort(key=lambda x: (x[2], x[3]), reverse=True)
        
        # Format for return (remove match count)
        self.detected_keys = [(col1, col2, score) for col1, col2, score, _ in candidates]
        
        logger.info(f"Detected {len(self.detected_keys)} potential merge key pairs")
        return self.detected_keys
    
    def validate_merge_keys(self, key1: str, key2: str) -> Dict[str, any]:
        """
        Validate selected merge keys and return statistics
        Args:
            key1: Column name in first DataFrame
            key2: Column name in second DataFrame
        Returns:
            Dictionary with validation results and statistics
        """
        if self.df1 is None or self.df2 is None:
            return {"valid": False, "error": "Files not loaded"}
        
        if key1 not in self.df1.columns:
            return {"valid": False, "error": f"Column '{key1}' not found in first file"}
            
        if key2 not in self.df2.columns:
            return {"valid": False, "error": f"Column '{key2}' not found in second file"}
        
        try:
            # Get clean values
            values1 = self.df1[key1].astype(str).str.strip()
            values2 = self.df2[key2].astype(str).str.strip()
            
            # Remove null/empty values for analysis
            clean_values1 = values1[values1.notna() & (values1 != '') & (values1 != 'nan')]
            clean_values2 = values2[values2.notna() & (values2 != '') & (values2 != 'nan')]
            
            # Calculate statistics
            unique1 = clean_values1.nunique()
            unique2 = clean_values2.nunique()
            total1 = len(clean_values1)
            total2 = len(clean_values2)
            
            # Find matches
            matches = set(clean_values1.str.upper()).intersection(set(clean_values2.str.upper()))
            
            result = {
                "valid": True,
                "file1_total": total1,
                "file1_unique": unique1,
                "file1_uniqueness": unique1 / total1 if total1 > 0 else 0,
                "file2_total": total2,
                "file2_unique": unique2,
                "file2_uniqueness": unique2 / total2 if total2 > 0 else 0,
                "common_values": len(matches),
                "match_ratio_file1": len(matches) / unique1 if unique1 > 0 else 0,
                "match_ratio_file2": len(matches) / unique2 if unique2 > 0 else 0,
            }
            
            # Add warnings
            warnings = []
            if result["file1_uniqueness"] < 0.7 and result["file2_uniqueness"] < 0.7:
                warnings.append("Both columns have low uniqueness - may not be good merge keys")
            elif result["match_ratio_file1"] < 0.3 and result["match_ratio_file2"] < 0.3:
                warnings.append("Low overlap between files - check if columns are compatible")
            
            result["warnings"] = warnings
            
            return result
            
        except Exception as e:
            return {"valid": False, "error": f"Error validating keys: {str(e)}"}
    
    def perform_merge(self, key1: str, key2: str, how: str = 'left') -> bool:
        """
        Perform the actual merge operation
        Args:
            key1: Column name in first DataFrame (left)
            key2: Column name in second DataFrame (right)  
            how: Type of merge ('left', 'right', 'inner', 'outer')
        Returns:
            Boolean indicating success
        """
        if self.df1 is None or self.df2 is None:
            logger.error("Both files must be loaded before merging")
            return False
        
        try:
            # Prepare DataFrames for merge
            left_df = self.df1.copy()
            right_df = self.df2.copy()
            
            # Normalize merge keys - FIXED: Simple version without regex issues
            left_df[key1] = left_df[key1].astype(str).str.strip().str.replace('.0', '')
            right_df[key2] = right_df[key2].astype(str).str.strip().str.replace('.0', '')
            
            # Check overlap
            left_keys = set(left_df[key1].unique())
            right_keys = set(right_df[key2].unique())
            common_keys = left_keys.intersection(right_keys)
            
            # Handle column name conflicts (except merge keys)
            right_columns = list(right_df.columns)
            
            if key2 != key1:
                right_df = right_df.rename(columns={key2: key1})
                right_columns[right_columns.index(key2)] = key1
                        
            # Rename conflicting columns in right DataFrame
            conflicts_renamed = {}
            for col in right_columns:
                if col in left_df.columns and col != key1:
                    new_name = f"{col}_right"
                    right_df = right_df.rename(columns={col: new_name})
                    conflicts_renamed[col] = new_name
                    print(f"Conflitto risolto: '{col}' -> '{new_name}'")
            
            
            # Data check before merge
            sample_right = right_df.head(3)
            for col in right_df.columns:
                if col != key1:  # Skip merge key
                    sample_values = sample_right[col].tolist()
                    non_null_count = right_df[col].notna().sum()
                    print(f"  {col}: {sample_values} (non-null: {non_null_count}/{len(right_df)})")
            print()

            # Perform merge
                        
            self.merge_result = pd.merge(
                left_df, 
                right_df, 
                on=key1, 
                how=how, 
                suffixes=('', '_right')
            )
            
                        # Checks columns 
            right_only_cols = [col for col in self.merge_result.columns if col not in left_df.columns or col.endswith('_right')]
            
            for col in right_only_cols[:5]:  
                if col in self.merge_result.columns:
                    non_null = self.merge_result[col].notna().sum()
                    total = len(self.merge_result)
                    if non_null > 0:
                        sample = self.merge_result[col].dropna().head(3).tolist()
                                
            logger.info(f"Merge completed: {len(self.merge_result)} rows in result")
            logger.info(f"Original files: {len(left_df)} + {len(right_df)} rows")
            
            return True
            
        except Exception as e:
            logger.error(f"Error performing merge: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def save_result(self, output_path: str) -> bool:
        """
        Save merge result to Excel file
        Args:
            output_path: Path where to save the result
        Returns:
            Boolean indicating success
        """
        if self.merge_result is None:
            logger.error("No merge result to save")
            return False
        
        try:
            # Determine output format from extension
            file_ext = os.path.splitext(output_path)[1].lower()
            
            if file_ext == '.xlsx':
                self.merge_result.to_excel(output_path, index=False)
            elif file_ext == '.csv':
                self.merge_result.to_csv(output_path, index=False)
            else:
                # Default to Excel
                if not output_path.endswith('.xlsx'):
                    output_path += '.xlsx'
                self.merge_result.to_excel(output_path, index=False)
            
            logger.info(f"Result saved to: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error saving result: {str(e)}")
            return False
    
    def get_preview_data(self, max_rows: int = 5) -> Dict[str, any]:
        """
        Get preview data for both loaded files
        Args:
            max_rows: Maximum number of rows to return for preview
        Returns:
            Dictionary with preview information
        """
        result = {}
        
        if self.df1 is not None:
            result['file1'] = {
                'name': os.path.basename(self.file1_path),
                'rows': len(self.df1),
                'columns': len(self.df1.columns),
                'column_names': list(self.df1.columns),
                'preview': self.df1.head(max_rows).to_dict('records')
            }
        
        if self.df2 is not None:
            result['file2'] = {
                'name': os.path.basename(self.file2_path), 
                'rows': len(self.df2),
                'columns': len(self.df2.columns),
                'column_names': list(self.df2.columns),
                'preview': self.df2.head(max_rows).to_dict('records')
            }
        
        return result


# Utility functions for data normalization
def normalize_tracking_code(value: str) -> str:
    """Normalize tracking codes by removing extra spaces and standardizing format"""
    if pd.isna(value) or value == '':
        return ''
    return str(value).strip().upper()

def detect_column_patterns(column_name: str) -> str:
    """Detect common column patterns for business context"""
    name_lower = column_name.lower().strip()
    
    tracking_patterns = ['tracking', 'track', 'awb', 'courier', 'shipment', 'spedizione']
    order_patterns = ['order', 'ordine', 'numero', 'reference', 'rif', 'comando']
    customer_patterns = ['customer', 'cliente', 'client', 'conto', 'account']
    status_patterns = ['status', 'stato', 'state', 'delivery', 'consegna']
    
    for pattern in tracking_patterns:
        if pattern in name_lower:
            return 'tracking'
    
    for pattern in order_patterns:
        if pattern in name_lower:
            return 'order'
    
    for pattern in customer_patterns:
        if pattern in name_lower:
            return 'customer'
    
    for pattern in status_patterns:
        if pattern in name_lower:
            return 'status'
    
    return 'unknown'