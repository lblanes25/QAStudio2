import os
import pandas as pd
import datetime
import json
import yaml
from typing import Dict, Optional, Union

from qa_analytics.utils.logging_config import setup_logging

logger = setup_logging()


class ReferenceDataManager:
    """
    Manages reference data with version control and freshness tracking
    """

    def __init__(self, config_path: str = "configs/reference_data.yaml"):
        """
        Initialize with reference data configuration

        Args:
            config_path: Path to reference data configuration file
        """
        self.config_path = config_path
        self.reference_data = {}
        self.metadata = {}
        self.config = {}
        self.audit_log = []

        # Load configuration
        self._load_config()

        # Initialize audit log
        self._init_audit_log()

    def _load_config(self) -> None:
        """Load reference data configuration"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = yaml.safe_load(f)
                logger.info(f"Loaded reference data config from {self.config_path}")
            else:
                logger.warning(f"Reference data config not found at {self.config_path}")
                self.config = {"reference_files": {}}
        except Exception as e:
            logger.error(f"Error loading reference data config: {e}")
            self.config = {"reference_files": {}}

    def _init_audit_log(self) -> None:
        """Initialize audit log for reference data changes"""
        log_path = self.config.get("audit_log_path", "logs/reference_data_audit.json")
        os.makedirs(os.path.dirname(log_path), exist_ok=True)

        if os.path.exists(log_path):
            try:
                with open(log_path, 'r') as f:
                    self.audit_log = json.load(f)
            except Exception as e:
                logger.error(f"Error loading audit log: {e}")
                self.audit_log = []

    def _save_audit_log(self) -> None:
        """Save audit log to file"""
        log_path = self.config.get("audit_log_path", "logs/reference_data_audit.json")
        try:
            with open(log_path, 'w') as f:
                json.dump(self.audit_log, f, indent=2, default=str)
        except Exception as e:
            logger.error(f"Error saving audit log: {e}")

    def load_reference_data(self, name: str, file_path: Optional[str] = None) -> bool:
        """
        Load reference data with freshness tracking

        Args:
            name: Reference data name
            file_path: Optional override for file path

        Returns:
            bool: True if loaded successfully
        """
        # Get config for this reference data
        if name not in self.config.get("reference_files", {}):
            logger.warning(f"No configuration found for reference data '{name}'")
            return False

        ref_config = self.config["reference_files"][name]

        # Use provided path or config path
        path = file_path or ref_config.get("path")
        if not path:
            logger.error(f"No file path specified for reference data '{name}'")
            return False

        if not os.path.exists(path):
            logger.error(f"Reference data file not found: {path}")
            return False

        try:
            # Get file modification time
            mod_time = os.path.getmtime(path)
            last_modified = datetime.datetime.fromtimestamp(mod_time)

            # Load the data based on file type
            file_ext = os.path.splitext(path)[1].lower()

            if file_ext == '.xlsx' or file_ext == '.xls':
                df = pd.read_excel(path)
            elif file_ext == '.csv':
                df = pd.read_csv(path)
            else:
                logger.error(f"Unsupported file type for reference data: {file_ext}")
                return False

            # Convert to dictionary if specified
            if ref_config.get("format") == "dictionary":
                key_column = ref_config.get("key_column")
                value_column = ref_config.get("value_column")

                if not key_column or not value_column:
                    logger.error(f"Missing key_column or value_column for dictionary reference data '{name}'")
                    return False

                # Create dictionary from key->value columns
                data = dict(zip(df[key_column], df[value_column]))
            else:
                # Use DataFrame as is
                data = df

            # Store data and metadata
            self.reference_data[name] = data
            version = ref_config.get("version", "1.0")

            self.metadata[name] = {
                'last_modified': last_modified,
                'file_path': path,
                'row_count': len(df),
                'columns': list(df.columns),
                'version': version,
                'loaded_at': datetime.datetime.now()
            }

            logger.info(f"Loaded reference data '{name}' version {version} from {path} ({len(df)} rows)")
            return True

        except Exception as e:
            logger.error(f"Error loading reference data '{name}': {e}")
            return False

    def get_reference_data(self, name: str) -> Union[Dict, pd.DataFrame, None]:
        """
        Get reference data by name

        Args:
            name: Reference data name

        Returns:
            Reference data (dict or DataFrame) or None if not found
        """
        # Check if already loaded
        if name in self.reference_data:
            return self.reference_data[name]

        # Try to load it
        if self.load_reference_data(name):
            return self.reference_data[name]

        return None

    def check_freshness(self, name: str, max_age_days: Optional[int] = None) -> bool:
        """
        Check if reference data is fresh

        Args:
            name: Reference data name
            max_age_days: Maximum age in days (overrides config)

        Returns:
            bool: True if fresh, False if stale or not found
        """
        if name not in self.metadata:
            return False

        # Get max age from parameter, config, or default
        if max_age_days is None:
            if name in self.config.get("reference_files", {}):
                max_age_days = self.config["reference_files"][name].get("max_age_days")

            if max_age_days is None:
                max_age_days = self.config.get("default_max_age_days", 30)

        last_modified = self.metadata[name]['last_modified']
        age = datetime.datetime.now() - last_modified

        return age.days <= max_age_days

    def get_freshness_status(self, name: str = None) -> Dict:
        """
        Get freshness status for all or one reference data

        Args:
            name: Optional reference data name

        Returns:
            Dict with freshness status information
        """
        if name:
            if name not in self.metadata:
                return {"status": "not_loaded", "message": "Reference data not loaded"}

            data_info = self.metadata[name]
            is_fresh = self.check_freshness(name)

            return {
                "name": name,
                "status": "fresh" if is_fresh else "stale",
                "last_modified": data_info['last_modified'],
                "age_days": (datetime.datetime.now() - data_info['last_modified']).days,
                "row_count": data_info['row_count'],
                "version": data_info['version']
            }
        else:
            # Return status for all reference data
            result = {}
            for ref_name in self.metadata:
                result[ref_name] = self.get_freshness_status(ref_name)
            return result

    def update_reference_data(self, name: str, file_path: str, user: str = "system") -> bool:
        """
        Update reference data with audit logging

        Args:
            name: Reference data name
            file_path: New file path
            user: User performing the update

        Returns:
            bool: True if update successful
        """
        # Store previous version info
        prev_version = None
        if name in self.metadata:
            prev_version = self.metadata[name].copy()

        # Load new data
        if not self.load_reference_data(name, file_path):
            return False

        # Log the change
        log_entry = {
            'timestamp': datetime.datetime.now(),
            'action': 'update_reference',
            'name': name,
            'user': user,
            'previous_version': prev_version,
            'new_version': self.metadata[name]
        }

        self.audit_log.append(log_entry)
        self._save_audit_log()

        # Log change
        if prev_version:
            logger.info(
                f"Reference data '{name}' updated by {user}: {prev_version['version']} -> {self.metadata[name]['version']}")
        else:
            logger.info(f"Reference data '{name}' initially loaded by {user}: {self.metadata[name]['version']}")

        return True

    def get_reference_data_info(self) -> Dict:
        """
        Get information about available reference data

        Returns:
            Dict with reference data information
        """
        info = {}

        # Include configured reference data
        for name, config in self.config.get("reference_files", {}).items():
            info[name] = {
                "name": name,
                "path": config.get("path", "Not specified"),
                "format": config.get("format", "dataframe"),
                "version": config.get("version", "Unknown"),
                "max_age_days": config.get("max_age_days", self.config.get("default_max_age_days", 30)),
                "description": config.get("description", ""),
                "loaded": name in self.reference_data
            }

            # Add metadata if loaded
            if name in self.metadata:
                info[name].update({
                    "last_modified": self.metadata[name]['last_modified'],
                    "row_count": self.metadata[name]['row_count'],
                    "loaded_at": self.metadata[name]['loaded_at'],
                    "is_fresh": self.check_freshness(name)
                })

        return info