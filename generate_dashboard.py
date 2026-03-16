"""
Excel to HTML Dashboard Generator
Reads Excel data and injects it into HTML template
"""

import pandas as pd
import json
from datetime import datetime
from pathlib import Path

# Paths
DATA_FILE = Path('data/Professional_Workload_Monthly capacity view.xlsx')
TEMPLATE_FILE = Path('docs/index.html')
OUTPUT_FILE = Path('docs/index.html')

def load_excel_data():
    """Load data from Excel sheets"""
    workload = pd.read_excel(DATA_FILE, sheet_name='Workload analysis')
    gantt = pd.read_excel(DATA_FILE, sheet_name='Gantt')
    
    return workload, gantt

def transform_to_raw(workload):
    """Transform Excel workload data to RAW array format"""
    raw_data = []
    
    for _, row in workload.iterrows():
        # Parse dates
        id_start = pd.to_datetime(row['ID Start Date']).strftime('%Y-%m-%d')
        id_end = pd.to_datetime(row['ID End Date']).strftime('%Y-%m-%d')
        sop_date = '2026-04-01'  # Default, can be enhanced
        
        raw_data.append({
            'project': row['Project'],
            'vertical': row['Vertical'],
            'type': row['Type'],
            'sop_date': sop_date,
            'id_start': id_start,
            'id_end': id_end,
            'duration': int(row['Duration']),
            'resource': row['Assigned Resource'],
            'role': row['Role']
        })
    
    return raw_data

def calculate_next_free(workload):
    """Calculate next free date for each resource"""
    resources = workload['Assigned Resource'].unique()
    next_free = []
    
    for resource in resources:
        resource_data = workload[workload['Assigned Resource'] == resource]
        max_end_date = pd.to_datetime(resource_data['ID End Date']).max()
        
        next_free.append({
            'resource': resource,
            'nextFreeDate': max_end_date.strftime('%d-%b-%y').upper(),
            'nextFreeMonth': max_end_date.strftime('%Y-%m'),
            'countEndingThatMonth': len(resource_data[pd.to_datetime(resource_data['ID End Date']).dt.strftime('%Y-%m') == max_end_date.strftime('%Y-%m')])
        })
    
    return sorted(next_free, key=lambda x: x['nextFreeDate'])

def get_unique_resources(workload):
    """Get list of unique resources"""
    return sorted(workload['Assigned Resource'].unique().tolist())

def generate_html(raw_data, next_free, resources):
    """Generate HTML with injected data"""
    
    # Read template
    with open(TEMPLATE_FILE, 'r', encoding='utf-8') as f:
        html_content = f.read()
    
    # Calculate metrics
    total_assignments = len(raw_data)
    unique_projects = len(set(r['project'] for r in raw_data))
    major_count = len([r for r in raw_data if r['type'] == 'Major'])
    minor_count = len([r for r in raw_data if r['type'] == 'Minor'])
    total_resources = len(resources)
    
    # Prepare JavaScript data
    raw_json = json.dumps(raw_data, indent=4)
    next_free_json = json.dumps(next_free, indent=4)
    resources_json = json.dumps(resources, indent=4)
    
    # Find and replace the RAW array in HTML
    import re
    
    # Replace RAW array
    raw_pattern = r'const RAW = \[.*?\];'
    html_content = re.sub(raw_pattern, f'const RAW = {raw_json};', html_content, flags=re.DOTALL)
    
    # Replace NEXT_FREE array
    nf_pattern = r'const NEXT_FREE = \[.*?\];'
    html_content = re.sub(nf_pattern, f'const NEXT_FREE = {next_free_json};', html_content, flags=re.DOTALL)
    
    # Replace RESOURCES array
    res_pattern = r'const RESOURCES = \[.*?\];'
    html_content = re.sub(res_pattern, f'const RESOURCES = {resources_json};', html_content, flags=re.DOTALL)
    
    # Update KPI values in HTML
    html_content = html_content.replace('"71 Assignments"', f'"{total_assignments} Assignments"')
    html_content = html_content.replace('"28 Projects"', f'"{unique_projects} Projects"')
    html_content = html_content.replace('"13 Resources"', f'"{total_resources} Resources"')
    html_content = html_content.replace('<div class="kpi-value">71</div>', f'<div class="kpi-value">{total_assignments}</div>')
    html_content = html_content.replace('<div class="kpi-value">28</div>', f'<div class="kpi-value">{unique_projects}</div>')
    html_content = html_content.replace('<div class="kpi-value">13</div>', f'<div class="kpi-value">{total_resources}</div>')
    html_content = html_content.replace('<div class="kpi-value">30</div>', f'<div class="kpi-value">{major_count}</div>')
    html_content = html_content.replace('<div class="kpi-value">41</div>', f'<div class="kpi-value">{minor_count}</div>')
    
    # Update date
    today = datetime.now().strftime('%b %d, %Y')
    html_content = html_content.replace('📅 Mar 07, 2026', f'📅 {today}')
    
    # Write output
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✅ Dashboard generated successfully!")
    print(f"   Total Assignments: {total_assignments}")
    print(f"   Unique Projects: {unique_projects}")
    print(f"   Resources: {total_resources}")
    print(f"   Output: {OUTPUT_FILE}")

def main():
    print("🔄 Starting dashboard generation...")
    
    # Load data
    workload, gantt = load_excel_data()
    print(f"📊 Loaded {len(workload)} workload records")
    
    # Transform data
    raw_data = transform_to_raw(workload)
    next_free = calculate_next_free(workload)
    resources = get_unique_resources(workload)
    
    # Generate HTML
    generate_html(raw_data, next_free, resources)
    
    print("✅ Done!")

if __name__ == '__main__':
    main()