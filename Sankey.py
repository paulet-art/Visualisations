import pandas as pd
import plotly.graph_objects as go
from collections import defaultdict
import numpy as np

# Extended WFP color palette
WFP_COLORS = [
    '#0A6EB4', '#7F1416', '#FFB400', '#007DBC', '#FF7F32', '#00833D', '#B4B4B4',
    '#2D9CDB', '#56CCF2', '#F2994A', '#EB5757', '#219653', '#6FCF97', '#BB6BD9',
    '#9B51E0', '#2F80ED'
]

def load_and_prepare_data(file_path):
    """Load and prepare the Excel data."""
    df = pd.read_excel(file_path)
    return df['Suggestions'].dropna().tolist()

def create_hierarchical_categories(suggestions):
    """Create detailed hierarchical categories from suggestions."""
    # Level 1: Main Categories
    main_categories = {
        'Communication & Engagement': ['communication', 'team', 'collaboration', 'meeting', 'sharing'],
        'Visibility & Advocacy': ['visibility', 'advocacy', 'branding', 'media', 'campaign'],
        'Data & Analytics': ['data', 'dashboard', 'analysis', 'metrics', 'reporting'],
        'Technology & Innovation': ['technology', 'automation', 'system', 'platform', 'AI'],
        'Capacity Building & Training': ['training', 'workshop', 'skill', 'learning', 'development'],
        'Partnerships & Stakeholders': ['partner', 'stakeholder', 'donor', 'collaboration', 'engagement'],
        'Process Improvement': ['efficiency', 'workflow', 'streamline', 'optimization', 'simplify']
    }

    # Level 2: Subcategories
    sub_categories = {
        'Communication & Engagement': {
            'Internal Collaboration': ['staff', 'unit', 'department', 'updates', 'internal'],
            'External Engagement': ['public', 'outreach', 'stakeholder', 'community'],
            'Information Sharing': ['newsletters', 'knowledge sharing', 'best practices']
        },
        'Visibility & Advocacy': {
            'Media Campaigns': ['press', 'promotion', 'storytelling', 'marketing'],
            'Branding & Recognition': ['logo', 'identity', 'branding', 'awareness'],
            'Public Relations': ['PR', 'messaging', 'reputation']
        },
        'Data & Analytics': {
            'Data Dashboards': ['dashboard', 'real-time', 'insights'],
            'Standardized Reporting': ['reporting', 'metrics', 'benchmarking'],
            'AI & Predictive Analytics': ['AI', 'machine learning', 'predictive']
        },
        'Technology & Innovation': {
            'Digital Tools': ['software', 'platform', 'tech solution'],
            'Process Automation': ['automation', 'workflow', 'streamlining'],
            'Cybersecurity & Data Protection': ['security', 'encryption', 'data protection']
        },
        'Capacity Building & Training': {
            'Workshops & Webinars': ['webinar', 'training session', 'courses'],
            'Knowledge Management': ['learning hub', 'resource center'],
            'Staff Development': ['career growth', 'upskilling']
        },
        'Partnerships & Stakeholders': {
            'Donor Relations': ['funding', 'donor engagement', 'sponsorship'],
            'Community Outreach': ['grassroots', 'volunteers', 'local partnerships'],
            'Corporate & NGO Partnerships': ['private sector', 'NGOs', 'government']
        },
        'Process Improvement': {
            'Efficiency & Optimization': ['cost-saving', 'process improvement'],
            'Workflow Automation': ['simplify', 'automate'],
            'Policy & Governance': ['compliance', 'best practices']
        }
    }

    hierarchy = defaultdict(lambda: defaultdict(list))

    for suggestion in suggestions:
        suggestion_lower = suggestion.lower()

        # Identify main category
        main_cat = 'Feedback'
        for category, keywords in main_categories.items():
            if any(keyword in suggestion_lower for keyword in keywords):
                main_cat = category
                break

        # Identify subcategory
        sub_cat = 'General'
        if main_cat in sub_categories:
            for subcategory, keywords in sub_categories[main_cat].items():
                if any(keyword in suggestion_lower for keyword in keywords):
                    sub_cat = subcategory
                    break

        hierarchy[main_cat][sub_cat].append(suggestion)

    return hierarchy

def create_sankey_data(hierarchy):
    """Create Sankey diagram data from hierarchy."""
    nodes = ['Suggestions']
    node_colors = []
    links = []
    values = []

    # Add main categories
    main_categories = list(hierarchy.keys())
    nodes.extend(main_categories)

    # Assign colors
    node_colors.append(WFP_COLORS[0])  # Root node color
    for i, _ in enumerate(main_categories):
        node_colors.append(WFP_COLORS[i % len(WFP_COLORS)])

    # Create links from root to main categories
    for main_cat in main_categories:
        links.append(('Suggestions', main_cat))
        values.append(sum(len(suggestions) for suggestions in hierarchy[main_cat].values()))

    # Add subcategories and their links
    for main_cat in main_categories:
        for sub_cat in hierarchy[main_cat].keys():
            node_name = f"{main_cat}-{sub_cat}"
            nodes.append(node_name)
            node_colors.append(WFP_COLORS[len(nodes) % len(WFP_COLORS)])
            links.append((main_cat, node_name))
            values.append(len(hierarchy[main_cat][sub_cat]))

    return nodes, links, values, node_colors

def create_sankey_diagram(nodes, links, values, node_colors, hierarchy):
    """Create an interactive Sankey diagram."""
    node_indices = {node: idx for idx, node in enumerate(nodes)}

    source = [node_indices[link[0]] for link in links]
    target = [node_indices[link[1]] for link in links]

    # Create hover text
    hover_text = []
    for node in nodes:
        if node == 'Suggestions':
            hover_text.append('All Suggestions')
        elif '-' not in node:  # Main category
            count = sum(len(suggestions) for suggestions in hierarchy[node].values())
            hover_text.append(f"{node}<br>Count: {count}")
        else:  # Sub-category
            main_cat, sub_cat = node.split('-')
            count = len(hierarchy[main_cat][sub_cat])
            suggestions = hierarchy[main_cat][sub_cat]
            hover_text.append(f"{sub_cat}<br>Count: {count}<br><br>Examples:<br>" +
                             "<br>".join(f"- {s[:100]}..." for s in suggestions[:5]))

    # Create figure
    fig = go.Figure(data=[go.Sankey(
        node=dict(
            pad=50,
            thickness=20,
            line=dict(color="black", width=0.5),
            label=[label[:70] + '...' if len(label) > 70 else label for label in nodes],
            color=node_colors,
            customdata=hover_text,
            hovertemplate='%{customdata}<extra></extra>'
        ),
        link=dict(
            source=source,
            target=target,
            value=values,
            hovertemplate='%{value} suggestions<extra></extra>'
        )
    )])

    # Update layout
    fig.update_layout(
        title_text="WFP Suggestions Analysis - Detailed Hierarchy",
        font_size=15,
        height=1000,
        paper_bgcolor='white',
        plot_bgcolor='white'
    )

    return fig

def main(file_path):
    # Load data
    suggestions = load_and_prepare_data(file_path)

    # Create hierarchy
    hierarchy = create_hierarchical_categories(suggestions)

    # Create Sankey data
    nodes, links, values, node_colors = create_sankey_data(hierarchy)

    # Create and display diagram
    fig = create_sankey_diagram(nodes, links, values, node_colors, hierarchy)
    fig.show()
    fig.write_html("sankey.html")


# Upload file in Colab before running
main('/content/Suggestions_Analysis.xlsx')
