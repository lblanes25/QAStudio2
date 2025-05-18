"""
merge_data_source.py - Utility to merge a generated data source into the data_sources.yaml file
"""

import os
import sys
import yaml
import argparse


def merge_data_source(input_yaml, output_yaml):
    """Merge a data source YAML into the main data sources configuration."""
    # Load input file
    with open(input_yaml, 'r') as f:
        new_config = yaml.safe_load(f)

    # Load existing config or create new
    existing_config = {}
    if os.path.exists(output_yaml):
        with open(output_yaml, 'r') as f:
            existing_config = yaml.safe_load(f) or {}

    # Make sure we have the right structure
    if 'data_sources' not in existing_config:
        existing_config['data_sources'] = {}

    if 'analytics_mapping' not in existing_config:
        existing_config['analytics_mapping'] = []

    # Find data source name (first key in the data_sources dict)
    if 'data_sources' not in new_config or not new_config['data_sources']:
        print("Error: No data sources found in input file")
        return False

    data_source_name = next(iter(new_config['data_sources'].keys()))

    # Merge the data source
    existing_config['data_sources'][data_source_name] = new_config['data_sources'][data_source_name]

    # Check if we need to add a new mapping
    mapping_exists = False
    for mapping in existing_config['analytics_mapping']:
        if mapping.get('data_source') == data_source_name:
            mapping_exists = True
            break

    if not mapping_exists:
        # Add from the new config if it exists
        for mapping in new_config.get('analytics_mapping', []):
            if mapping.get('data_source') == data_source_name:
                existing_config['analytics_mapping'].append(mapping)
                break
        else:
            # Add empty mapping if not found
            existing_config['analytics_mapping'].append({
                'data_source': data_source_name,
                'analytics': []
            })

    # Write the merged config
    with open(output_yaml, 'w') as f:
        yaml.dump(existing_config, f, default_flow_style=False, sort_keys=False)

    print(f"Successfully merged '{data_source_name}' into {output_yaml}")
    return True


def main():
    parser = argparse.ArgumentParser(description="Merge a generated data source into data_sources.yaml")
    parser.add_argument("input_file", help="Generated data source YAML file")
    parser.add_argument("--output", default="configs/data_sources.yaml",
                        help="Path to data sources config (default: configs/data_sources.yaml)")

    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        print(f"Error: Input file not found: {args.input_file}")
        return 1

    # Create output directory if needed
    output_dir = os.path.dirname(args.output)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)

    if merge_data_source(args.input_file, args.output):
        return 0
    else:
        return 1


if __name__ == "__main__":
    sys.exit(main())