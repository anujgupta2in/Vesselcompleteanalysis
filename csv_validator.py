import pandas as pd
from datetime import datetime

class CSVValidator:
    def __init__(self):
        # Define required columns and their possible alternative names
        self.column_mappings = {
            'Vessel': ['Vessel', 'Ship Name', 'Vessel Name'],
            'Job Code': ['Job Code', 'CMS Code', 'Work Code', 'Task Code'],
            'Frequency': ['Frequency', 'Interval', 'Maintenance Frequency'],
            'Calculated Due Date': ['Calculated Due Date', 'Due Date', 'Next Due Date'],
            'Machinery Location': ['Machinery Location','Machinery', 'Equipment Location', 'Machine Location'],
            'Sub Component Location': ['Sub Component Location','Component', 'Component Location', 'Part Location']
        }
        self.required_columns = list(self.column_mappings.keys())
        self.numeric_columns = ['Job Code']

    def _find_matching_column(self, columns, possible_names):
        """Find a matching column name from possible alternatives."""
        for name in possible_names:
            if name in columns:
                return name
        return None

    def _suggest_column_mappings(self, df):
        """Suggest possible column mappings for missing required columns."""
        suggestions = {}
        available_columns = df.columns.tolist()
        for required_col, alternatives in self.column_mappings.items():
            if required_col not in df.columns:
                matched_col = self._find_matching_column(available_columns, alternatives)
                if matched_col:
                    suggestions[required_col] = matched_col
        return suggestions

    def validate_columns(self, df):
        """Check if all required columns are present and suggest mappings."""
        missing_columns = []
        suggestions = self._suggest_column_mappings(df)
        for col in self.required_columns:
            if col not in df.columns and col not in suggestions:
                missing_columns.append(col)
        if missing_columns:
            error_msg = "Missing required columns:\n"
            for col in missing_columns:
                error_msg += f"- {col}\n"
            if suggestions:
                error_msg += "\nPossible column mappings found:\n"
                for req_col, matched_col in suggestions.items():
                    error_msg += f"- '{matched_col}' could be used for '{req_col}'\n"
            return False, error_msg
        return True, ""

    def validate_numeric_fields(self, df):
        """Validate numeric fields and provide detailed error messages."""
        errors = []
        for col in self.numeric_columns:
            if col not in df.columns:
                continue
            numeric_data = pd.to_numeric(df[col], errors='coerce')
            non_numeric = df[numeric_data.isna() & df[col].notna()]
            if not non_numeric.empty:
                rows = non_numeric.index.tolist()
                total_rows = len(non_numeric)
                if total_rows > 0:
                    errors.append(f"Non-numeric values found in {col} at rows: {rows}")
        return len(errors) == 0, "\n".join(errors)

    def validate_machinery_location(self, df):
        """Validate machinery location format with detailed error reporting."""
        if 'Machinery Location' not in df.columns:
            return False, "Missing 'Machinery Location' column"
        # Filter only engine-related entries
        engine_entries = df[df['Machinery Location'].str.contains('Main Engine|Auxiliary Engine', case=False, na=False)]
        valid_patterns = [
            r'^Main Engine[\s-]*#?\d+$',
            r'^Main Engine[\s-]*MC[\s-]*#?\d+$',
            r'^Main Engine[\s-]*ME[\s-]*C[\s-]*II[\s-]*#?\d+$',
            r'^Main Engine[\s-]*ME[\s-]*C[\s-]*GI[\s-]*#?\d+$',
            r'^Main Engine[\s-]*ME[\s-]*C[\s-]*#?\d+$',
            r'^Main Engine[\s-]*ME[\s-]*B[\s-]*#?\d+$',
            r'^Main Engine[\s-]*RT[\s-]*FLEX[\s-]*#?\d+$',
            r'^Main Engine[\s-]*RTFLEX[\s-]*#?\d+$',
            r'^Main Engine[\s-]*RTA[\s-]*#?\d+$',
            r'^Main Engine[\s-]*UEC[\s-]*#?\d+$',
            r'^Main Engine[\s-]*W[\s-]*#?\d+$',
            r'^Main Engine[\s-]*WX[\s-]*#?\d+$',
            r'^Main Engine[\s-]*No\d+$',
            
            r'^Auxiliary Engine[\s-]*#?\d+$',
            r'^Auxiliary Engine[\s-]*No\d+$',
        ]


      
        import re
        pattern = '|'.join(valid_patterns)
        
        # Check for standard patterns first
        mask = engine_entries['Machinery Location'].apply(lambda x: bool(re.match(pattern, str(x))) if pd.notna(x) else False)
        
        # For entries not matching standard patterns, attempt auto-correction
        invalid_entries = engine_entries[~mask].copy()
        if not invalid_entries.empty:
            # Add a new column with corrected format (for potential fixes)
            df['_machinery_location_fixed'] = df['Machinery Location']
            # Fix patterns like "Auxiliary EngineNo4" to "Auxiliary Engine#4"
            df.loc[df['Machinery Location'].str.contains('No\\d+', regex=True, na=False), '_machinery_location_fixed'] = \
                df.loc[df['Machinery Location'].str.contains('No\\d+', regex=True, na=False), 'Machinery Location'].str.replace('No', '#', regex=False)
        
        # Original mask for validation feedback
        invalid_locations = engine_entries[~mask]
        if not invalid_locations.empty:
            sample_locations = invalid_locations['Machinery Location'].head().tolist()
            rows = invalid_locations.index.tolist()[:5]
            total_rows = len(invalid_locations)
            error_msg = (
                f"Invalid Engine Location format at rows: {rows}\n"
                f"Example invalid formats: {sample_locations}\n"
                "Expected format for engine entries:\n"
                "- Main Engine#X\n"
                "- Main Engine - MC#X\n"
                "- Main Engine - ME-C II#X\n"
                "- Auxiliary Engine#X\n"
                "where X is a number"
            )
            if total_rows > 5:
                error_msg += f"\nAnd {total_rows - 5} more rows contain invalid formats"
            return False, error_msg
        return True, ""

    def validate_data(self, df):
        """Perform all validations with comprehensive error reporting."""
        try:
            validations = [
                (self.validate_columns(df), "Column Validation"),
                (self.validate_numeric_fields(df), "Numeric Fields Validation"),
                (self.validate_machinery_location(df), "Engine Location Validation")
            ]
            errors = []
            for (is_valid, error_msg), category in validations:
                if not is_valid:
                    errors.append(f"\n{category}:\n{error_msg}")
            return len(errors) == 0, errors
        except Exception as e:
            return False, [f"Validation error: {str(e)}"]
