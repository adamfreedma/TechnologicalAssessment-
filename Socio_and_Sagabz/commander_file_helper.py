from re import split
from typing import List
import matplotlib.pyplot as plt
import constants
import numpy as np
import pandas as pd

avg_person = "total"

def generate_all_graphs(averages, stds):
    file_path = "Excels/data.xlsx"
    df = pd.read_excel(file_path)
    
    
    
    # Generate the average comparison graph
    average_compare_graph(averages, stds, categories=constants.PROFESSIONAL_CATEGORIES)
    average_compare_graph(averages, stds, categories=constants.PERSONAL_CATEGORIES)

    # Generate the group comparison graphs
    generate_groups(df, averages, categories=constants.PROFESSIONAL_CATEGORIES)
    generate_groups(df, averages, categories=constants.PERSONAL_CATEGORIES)
    
    # Generate the correlation heatmap
    data = pd.read_excel("Excels/combined_data_2.xlsx")
    correlation_heatmap(data)


def generate_groups(df, averages, categories=constants.PROFESSIONAL_CATEGORIES):
        
    all_names = df['name'].tolist()
    all_names_set = set(all_names)  # Convert to a set for unique names

    # create names of groups
    male_names = df.loc[df['gender'] == 'זכר', 'name'].tolist()
    female_names = list(all_names_set - set(male_names))
    
    split_bar_graphs([male_names, female_names], ["גברים", "נשים"], averages, "גרף השוואה בין מגדרים", categories)
    
    
    boofor_names = df.loc[df['platoon'] == 'בופור', 'name'].tolist()
    reim_names = df.loc[df['platoon'] == 'רעים', 'name'].tolist()
    magen_names = df.loc[df['platoon'] == 'מגן', 'name'].tolist()
    sufa_names = df.loc[df['platoon'] == 'סופה', 'name'].tolist()
    
    split_bar_graphs([boofor_names, reim_names, magen_names, sufa_names], ["בופור", "רעים", "מגן", "סופה"], averages, "גרף השוואה בין מחלקות", categories)
    
    religous_names = df.loc[df['religious'] == 'דתי', 'name'].tolist()
    unreligous_names = list(all_names_set - set(religous_names))
    
    split_bar_graphs([religous_names, unreligous_names], ["דתיים", "חילוניים"], averages, "גרף השוואה בין דתיים לחילוניים", categories)
    
    physics_names = df.loc[df['major'] == 'פיסיקה', 'name'].tolist()
    math_names = df.loc[df['major'] == 'מתמטיקה', 'name'].tolist()
    computer_names = df.loc[df['major'] == 'מדעי המחשב', 'name'].tolist()
    
    split_bar_graphs([physics_names, math_names, computer_names], ["פיזיקה", "מתמטיקה", "מדעי המחשב"], averages, "גרף השוואה בין מסלולים", categories)
    
    sterile_names = df.loc[df['background'] == 'סטרילי', 'name'].tolist()
    light_proritorist_names = df.loc[df['background'] == 'פטוריסט קל', 'name'].tolist()
    heavy_prtorist_names = df.loc[df['background'] == 'פטוריסט כבד', 'name'].tolist()
    degree_names = df.loc[df['background'] == 'תאריסט', 'name'].tolist()
    
    split_bar_graphs([sterile_names, light_proritorist_names, heavy_prtorist_names, degree_names], ["סטרילי", "פטוריסט קל", "פטוריסט כבד", "תאריסט"], averages, "גרף השוואה בין רקעים אקדמיים", categories)
    

def average_compare_graph(averages, stds, categories=constants.PROFESSIONAL_CATEGORIES):
    fig, ax = plt.subplots()

    # Calculate averages and standard deviations for each category
    avg_values = [averages[avg_person][-1].get(category, 0) for category in categories]
    std_values = [stds[avg_person][-1].get(category, 0) for category in categories]

    # Plot settings with error bars
    y_pos = np.arange(len(categories))
    bars = ax.barh(y_pos, avg_values, xerr=std_values, align='center', color='blue', edgecolor='black', alpha=0.6, label='ממוצע מחזורי'[::-1])

    # Labels and formatting
    ax.set_yticks(y_pos)
    if categories == constants.PROFESSIONAL_CATEGORIES:
        ax.set_xlim(1, 6)
    else:
        ax.set_xlim(1, 3)
    ax.set_yticklabels([constants.CATEGORY_NAME_DICT[categories[i]][::-1] for i in range(len(categories))], fontsize=constants.SMALL_FONTSIZE)
    ax.tick_params(axis='x', labelsize=constants.SMALL_FONTSIZE)

    # Add a legend for both average and standard deviation
    avg_patch = plt.Line2D([0], [0], color='blue', lw=4, label='ממוצע מחזורי'[::-1])
    std_patch = plt.Line2D([0], [0], color='black', lw=1, label='סטיית תקן'[::-1])
    ax.legend(handles=[avg_patch, std_patch], fontsize=constants.SMALL_FONTSIZE, loc='upper left')

    plt.gca().invert_yaxis()  # Invert y-axis to match typical bar chart order

    # Title and layout
    plt.title("ממוצעים מחזוריים"[::-1], fontsize=constants.SMALL_FONTSIZE)
    plt.tight_layout()
    plt.show()

    # Add a second graph with an additional bar for the overall average
    last_avg_values = [averages[avg_person][-2].get(category, 0) for category in categories]

    # Create a new figure for the second graph
    fig, ax = plt.subplots()

    # Add the overall average bar
    ax.barh(y_pos, avg_values, align='center', color='blue', edgecolor='black', alpha=0.6, label='ממוצע מחזורי'[::-1])
    ax.barh(y_pos, last_avg_values, align='center', color='green', edgecolor='black', alpha=0.6, label='ממוצע קודם'[::-1])

    # Update the legend to include both average and overall average
    avg_patch = plt.Line2D([0], [0], color='darkgreen', lw=4, label='ממוצע מחזורי'[::-1])
    overall_patch = plt.Line2D([0], [0], color='lightgreen', lw=4, label='ממוצע קודם'[::-1])
    ax.legend(handles=[avg_patch, overall_patch], fontsize=constants.SMALL_FONTSIZE, loc='upper left')

    # Update the title to reflect the addition of the overall average
    ax.set_yticks(y_pos)
    if categories == constants.PROFESSIONAL_CATEGORIES:
        ax.set_xlim(1, 6)
    else:
        ax.set_xlim(1, 3)
    ax.set_yticklabels([constants.CATEGORY_NAME_DICT[categories[i]][::-1] for i in range(len(categories))], fontsize=constants.SMALL_FONTSIZE)
    ax.tick_params(axis='x', labelsize=constants.SMALL_FONTSIZE)
    plt.title("ממוצעים מחזוריים בהשוואה לסמסטר קודם"[::-1], fontsize=constants.SMALL_FONTSIZE)
    plt.tight_layout()
    plt.show()

def split_bar_graphs(groups: List[List[str]], names: List[str], averages, title: str, categories=constants.PROFESSIONAL_CATEGORIES):
    fig, ax = plt.subplots(figsize=(10, 7))

    # Calculate averages and standard deviations for each group
    group_avg_values = [
        [np.mean([averages[g][-1].get(category, 0) for g in group]) for category in categories]
        for group in groups
    ]
    group_std_values = [
        [np.std([averages[g][-1].get(category, 0) for g in group]) for category in categories]
        for group in groups
    ]

    # Bar width and positions
    bar_width = 0.55 - 0.1 * len(groups)  # Adjust bar width based on number of groups
    x_pos = np.arange(len(categories))

    # Plot bars for each group
    for i, (avg_values, std_values) in enumerate(zip(group_avg_values, group_std_values)):
        ax.bar(
            x_pos + i * bar_width,
            avg_values,
            yerr=std_values,
            width=bar_width,
            align='center',
            label=names[i][::-1],
            alpha=0.6,
            edgecolor='black'
        )

    # Labels and formatting
    ax.set_xticks(x_pos + bar_width / 2)
    ax.set_xticklabels(
        [constants.CATEGORY_NAME_DICT[categories[j]][::-1] for j in range(len(categories))],
        fontsize=constants.SMALL_FONTSIZE
    )
    ax.tick_params(axis='y', labelsize=constants.SMALL_FONTSIZE)
    ax.legend(fontsize=constants.SMALL_FONTSIZE)

    # Title and layout
    ax.set_title(title[::-1], fontsize=constants.SMALL_FONTSIZE)
    plt.tight_layout()
    plt.show()
    
    
def pair_correlation(df, col1, col2):
    """
    Calculate the correlation between two columns in the dataframe.
    Exactly like df.describe does.
    """
    # Find the indexes where both columns are not NaN
    indexes = df[col1].notna() & df[col2].notna()
    # Get the columns
    col1 = df[col1][indexes]
    col2 = df[col2][indexes]
    # Calculate the correlation
    correlation = np.corrcoef(col1, col2)[0, 1]
    return correlation

def correlation_heatmap(df):
    
    cols = constants.TABLE_CATEGORIES
    cols_names = [constants.CATEGORY_NAME_DICT[x][::-1] for x in cols]
    corr_arr = np.zeros((len(cols), len(cols)))

    
        # Calculate the correlation
    for i in range(len(cols)):
        for j in range(i, len(cols)):
            col1 = cols[i]
            col2 = cols[j]
            cor = pair_correlation(df, col1, col2)
            corr_arr[i, j] = cor
            corr_arr[j, i] = cor

    # Plot the correlation between the columns as heatmap
    plt.figure(figsize=(8, 6))
    plt.title('טבלת קורצלציה'[::-1])
    plt.imshow(corr_arr, cmap="summer", interpolation='nearest')
    plt.colorbar()
    plt.xticks(range(len(corr_arr)), cols_names, rotation=45, )
    plt.yticks(range(len(corr_arr)), cols_names)

    # Insert numbers in the heatmap
    for i in range(len(corr_arr)):
        for j in range(len(corr_arr)):
            plt.text(j, i, f"{corr_arr[i, j]:.2f}", ha='center', va='center', color='black')

    plt.tight_layout()
    plt.show()