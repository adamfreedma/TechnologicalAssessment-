from abc import ABC, abstractmethod
from cProfile import label
from io import BytesIO
from itertools import count
from math import e
from operator import imod, is_, le
from tempfile import template

import docx.table
import docx.text
import docx.text.paragraph
import matplotlib.pyplot as plt
import matplotlib

import commander_file_helper

matplotlib.rcParams["font.family"] = "David"  # Set the font globally
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from matplotlib.figure import Figure
from tqdm import tqdm
import os
from docx.oxml.ns import qn
from column_constants import SIGMAS
from docxtpl import DocxTemplate
from docx.oxml import OxmlElement
from docx.shared import RGBColor
import docx
import constants


ADD_IN_END_OF_SENTENCE: str = "."
IS_SOCIO = True


def add_labels(ax, bars):
    for bar in bars:
        height = bar.get_height()
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            height,
            f"{height:.2f}",
            ha="center",
            va="bottom",
            fontsize=constants.FONTSIZE,
            color="black",
        )


def text_to_rgba(s, *, dpi, **kwargs):
    # To convert a text string to an image, we can:
    # - draw it on an empty and transparent figure;
    # - save the figure to a temporary buffer using ``bbox_inches="tight",
    #   pad_inches=0`` which will pick the correct area to save;
    # - load the buffer using ``plt.imread``.
    #
    # (If desired, one can also directly save the image to the filesystem.)
    fig = Figure(facecolor="none")
    fig.text(0, 0, s, **kwargs)
    with BytesIO() as buf:
        fig.savefig(buf, dpi=dpi, format="png", bbox_inches="tight", pad_inches=0)
        buf.seek(0)
        rgba = plt.imread(buf)
    return rgba


def fix_rtl_symbols(sentence):
    """Fix punctuation in RTL text using BiDi marks."""
    if not isinstance(sentence, str):
        return sentence

    # Add RLM around punctuation
    rtl_marks = "\u200f"
    sentence = sentence.replace(",", f"{rtl_marks},{rtl_marks}")
    sentence = sentence.replace(".", f"{rtl_marks}.{rtl_marks}")
    sentence = sentence.replace('"', f'{rtl_marks}"{rtl_marks}')
    sentence = sentence.replace("\\", f"{rtl_marks}\\{rtl_marks}")
    sentence = sentence.replace("-", f"{rtl_marks}-{rtl_marks}")
    sentence = sentence.replace(":", f"{rtl_marks}:{rtl_marks}")
    sentence = sentence.replace(";", f"{rtl_marks};{rtl_marks}")
    sentence = sentence.replace("(", f"{rtl_marks}({rtl_marks}")
    sentence = sentence.replace(")", f"{rtl_marks}){rtl_marks}")

    # Wrap entire sentence with RLE and PDF
    sentence = f"\u202b{sentence}\u202c"
    return sentence


def add_in_the_beginning(sentence, add_on):
    if isinstance(sentence, str) and sentence[0] != add_on:
        sentence = add_on + sentence
    return sentence


def ensure_ends_with(sentence, add_on):
    if isinstance(sentence, str):
        if len(sentence) == 0:
            pass
        elif sentence[-1] == " ":
            sentence = sentence[:-1]
        elif sentence[-1] != add_on:
            sentence += add_on
    sentence = fix_rtl_symbols(sentence)
    return sentence


class Docx_helper(ABC):
    def __init__(self, file_format_path, word_output_dir):
        self.file_format_path = file_format_path
        self.word_output_dir = word_output_dir

    @abstractmethod
    def is_values(self, category: str) -> bool:
        pass

    @abstractmethod
    def get_columns_names(self):
        pass

    def my_hash(self, s):
        return abs(hash(s)).__str__()

    def find_table_cell_by_tag(self, doc: Document, tag: str) -> docx.table._Cell:
        """finds a table cell in the document by a specific tag.

        Args:
            doc (Document): document object to search in
            tag (str): tag to search for in the table cells

        Returns:
            docx.table._Cell: a cell object if found, None otherwise
        """
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if tag in cell.text:
                        return cell

    def find_paragraph_by_tag(
        self, doc: Document, tag: str
    ) -> docx.text.paragraph.Paragraph:
        """find a paragraph in the document by a specific tag.

        Args:
            doc (Document): document object to search in
            tag (str): tag to search for in the paragraphs

        Returns:
            docx.text.paragraph.Paragraph: a paragraph object if found, None otherwise
        """
        for paragraph in doc.paragraphs:
            if tag in paragraph.text:
                return paragraph

    def fill_row(
        self,
        doc: Document,
        values,
        type,
        categories=constants.TABLE_CATEGORIES,
        color=RGBColor(0, 0, 0),
    ) -> None:

        if "commander" in categories:
            categories.remove("commander")

        if type in ["old ", "new ", "std ", "total ", "range ", "N "]:
            categories.append("commander")
            categories = set(categories)

        for category in categories:

            cell = self.find_table_cell_by_tag(doc, type + category)

            if cell is not None:
                # fill the cell with the data
                text = "".join(run.text for run in cell.paragraphs[0].runs)

                split = text.split("{")

                if len(split) != 2:
                    raise ValueError(
                        f"Cell with tag {type + category} contains a non 1 number of tags."
                    )

                before, after_with_tag = split
                after = after_with_tag.split("}")[1]
                cell.paragraphs[0].clear()
                cell.paragraphs[0].add_run(before)

                if category == "commander":
                    pass

                if category in values and values[category] is not None:
                    if isinstance(values[category], tuple):
                        cell.paragraphs[0].add_run(
                            f"{values[category][0]:.2f}-{values[category][1]:.2f}"
                        )
                    else:
                        cell.paragraphs[0].add_run(
                            f"{values[category]:.2f}".rstrip("0").rstrip(".")
                        )

                else:
                    cell.paragraphs[0].add_run(" - ")

                cell.paragraphs[0].add_run(after)

                for run in cell.paragraphs[0].runs:
                    run.font.name = "David"
                    run.font.size = Pt(12)
                    run.bold = True
                    if run.text != "0":
                        run.font.color.rgb = color
            else:
                exit(f"Cell with tag {type + category} not found in the document.")

    def fill_main_table(
        self, doc: Document, averages, stds, counts, person_name
    ) -> None:
        """Fill the main table in the document with the averages, stds, and counts for a specific person."""

        # find the table in the document and fill it with data
        old_data = {} if len(averages[person_name]) == 1 else averages[person_name][-2]
        new_data = averages[person_name][-1]
        stds_data = stds[person_name][-1]
        total_data = averages["total"][-1]
        range_data = averages["range"][-1]
        count_data = counts[person_name][-1]["total"]

        # fill the table with the data
        self.fill_row(doc, old_data, "old ")
        self.fill_row(doc, new_data, "new ")
        self.fill_row(doc, stds_data, "std ")
        self.fill_row(doc, total_data, "total ")
        self.fill_row(doc, range_data, "range ")
        self.fill_row(doc, count_data, "N ")

    def fill_values_table(self, doc: Document, counts, person_name: str) -> None:
        """Fill the values table in the document with the counts for a specific person."""
        name_dict = {1: "negative ", 2: "neutral ", 3: "positive "}
        color_dict = {
            1: constants.RED_COLOR,
            2: RGBColor(0, 0, 0),
            3: constants.GREEN_COLOR,
        }

        for value in range(1, 4):
            if value not in counts[person_name][-1]:
                exit(f"Value {value} not found in counts for {person_name}. Skipping.")
                continue

            # fill the table with the data
            self.fill_row(
                doc,
                counts[person_name][-1][value],
                name_dict[value],
                categories=constants.PERSONAL_CATEGORIES,
                color=color_dict[value],
            )

    def fill_semester_table(self, doc: Document, averages, person_name: str) -> None:
        """Fill the semester table in the document with the averages for a specific person."""

        for i in range(len(averages[person_name])):
            self.fill_row(
                doc,
                averages[person_name][i],
                f"{i+1} ",
                categories=constants.TABLE_CATEGORIES,
            )

    def create_main_graph(
        self,
        doc,
        averages,
        stds,
        person_name,
        path_to_save,
        tag,
        categories=constants.MAIN_CATEGORIES,
        scale=6,
        title="",
    ) -> None:
        # Plot settings
        fig, ax = plt.subplots(figsize=(23, 12))

        y_pos = np.arange(len(categories))

        avg_values = [
            averages[person_name][-1].get(category, 0) for category in categories
        ]
        total_avg_values = [
            averages["total"][-1].get(category, 0) for category in categories
        ]
        std_values = [stds[person_name][-1].get(category, 0) for category in categories]

        # plot the bars
        ax.barh(
            y_pos,
            total_avg_values,
            align="center",
            color="skyblue",
            edgecolor="black",
            label="ממוצע מחזורי"[::-1],
        )
        ax.errorbar(
            avg_values,
            y_pos,
            xerr=std_values,
            fmt="o",
            color="blue",
            label="ממוצע + פיזור"[::-1],
        )

        # Labels and formatting
        ax.set_yticks(y_pos)
        ax.set_yticklabels(
            [
                constants.CATEGORY_NAME_DICT[categories[i]][::-1]
                for i in range(len(categories))
            ],
            fontsize=constants.FONTSIZE,
        )
        ax.tick_params(axis="x", labelsize=constants.FONTSIZE)
        ax.legend(fontsize=constants.FONTSIZE)
        plt.gca().invert_yaxis()  # Invert y-axis to match typical bar chart order

        # save the figure
        plt.xlim(1, scale)
        plt.title(title[::-1], fontsize=constants.FONTSIZE)
        plt.savefig(path_to_save, bbox_inches="tight")
        plt.close(fig)

        self.add_graph(doc, path_to_save, tag) # MARK IMPORTANT

    def add_graph(self, doc, path_to_save, tag) -> None:
        """Add a graph to the document in a specific cell identified by a tag."""
        if doc is None:
            print(f"Error: 'doc' object is None. Cannot add graph for tag: {tag}")
            return

        cell = self.find_table_cell_by_tag(doc, tag)

        if cell is None:
            exit(f"Cell with tag {tag} not found in the document.")

        # Clear the cell and add the image
        cell.paragraphs[0].clear()
        if tag == "knowing_graph":
            cell.add_paragraph().add_run().add_picture(path_to_save, height=Inches(3.5))
        else:
            cell.add_paragraph().add_run().add_picture(path_to_save, height=Inches(2.5))
        cell.paragraphs[1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # remove png as it is no longer needed
        # os.remove(path_to_save)

    def create_knowing_graph(
        self, doc, high_knowing, low_knowing, person_name, path_to_save, tag
    ) -> None:
        bar_width = 0.4  # 0.25 and 0.1
        spacing = 0.15  # Extra spacing between groups
        x = np.arange(len(constants.KNOWING_CATEGORIES)) * (1 + spacing)

        # Create the plot
        fig, ax = plt.subplots(figsize=(40, 30))
        high_knowing_list = [
            high_knowing[person_name][-1][category]
            for category in constants.KNOWING_CATEGORIES
        ]
        low_knowing_list = [
            low_knowing[person_name][-1][category]
            for category in constants.KNOWING_CATEGORIES
        ]
        bars1 = ax.bar(
            x - bar_width,
            high_knowing_list,
            bar_width,
            label="מידת היכרות גבוהה < 4"[::-1],
            color="#ADD8E6",
        )  # Light blue
        bars2 = ax.bar(
            x,
            low_knowing_list,
            bar_width,
            label="מידת היכרות נמוכה >= 4"[::-1],
            color="#6495ED",
        )  # Medium blue

        add_labels(ax, bars1)
        add_labels(ax, bars2)

        ax.set_xticks(x)
        labels = [
            constants.CATEGORY_NAME_DICT[category][::-1]
            for category in constants.KNOWING_CATEGORIES
        ]
        ax.set_xticklabels(
            labels, rotation=45, ha="right", fontsize=2 * constants.FONTSIZE
        )  # Align to right for Hebrew
        ax.tick_params(axis="y", labelsize=2 * constants.FONTSIZE)
        ax.legend(fontsize=2 * constants.FONTSIZE)

        # Show the plot
        plt.title(
            "גרף היכרות, מדדים מפוצלים למידת היכרות נמוכה וגבוהה"[::-1] + "\n",
            fontsize=2 * constants.FONTSIZE,
        )
        plt.tight_layout()
        plt.ylim(1, 6)
        plt.savefig(path_to_save, bbox_inches="tight")
        plt.close(fig)

        self.add_graph(doc, path_to_save, tag) # MARK IMPORTANT

    def create_progress_graph(
        self, doc, averages, person_name, path_to_save, tag
    ) -> None:
        """Create a semester progress graph for a specific person and add it to the document."""
        COLORS = ["#ADD8E6", "#6495ED", "#00008B"]  # Light blue, Medium blue, Dark blue
        # Set bar width and spacing
        bar_width = 0.25
        spacing = 0.35  # Extra spacing between groups

        x = np.arange(2) * (1 + spacing)
        # Create the plot
        fig, ax = plt.subplots(figsize=(24, 10))

        semester_count = len(averages[person_name])
        for i in range(semester_count):

            # Calculate the average values for commandership and professionalism
            commandership = np.average(
                [
                    averages[person_name][i][category]
                    for category in constants.COMMANDERSHIP_CATEGORIES
                ]
            )
            professionalism = np.average(
                [
                    averages[person_name][i][category]
                    for category in constants.PROFESSIONAL_GRAPH_CATEGORIES
                ]
            )

            # Create bars for each semester
            bars = ax.bar(
                x + i * bar_width,
                [commandership, professionalism],
                bar_width,
                label=f"סמסטר {chr(ord('א') + i)}"[::-1],
                color=COLORS[i % len(COLORS)],
            )

            # Add labels to the bars
            for bar in bars:
                height = bar.get_height()
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    height,
                    f"{height:.2f}",
                    ha="center",
                    va="bottom",
                    fontsize=constants.FONTSIZE,
                    color="black",
                )

        ax.set_xticks(x + bar_width * (semester_count - 1) / 2)
        ax.set_xticklabels(
            ["מנהיגות"[::-1], "מקצועיות"[::-1]],
            rotation=45,
            ha="right",
            fontsize=constants.FONTSIZE,
        )  # Align to right for Hebrew
        ax.tick_params(axis="y", labelsize=constants.FONTSIZE)
        ax.legend(fontsize=constants.FONTSIZE)
        plt.tight_layout()
        plt.title("ממוצעים לאורך סמסטרים"[::-1], fontsize=constants.FONTSIZE)
        plt.ylim(1, 6.3)
        plt.savefig(path_to_save, bbox_inches="tight")
        plt.close(fig)

        self.add_graph(doc, path_to_save, tag)

    def set_title(self, doc: Document, title: str) -> None:
        """Set the title of the document."""
        paragraph = self.find_paragraph_by_tag(doc, "title")
        paragraph.clear()
        paragraph.add_run(title)
        paragraph.runs[0].bold = True
        paragraph.runs[0].underline = True
        paragraph.runs[0].font.name = "David"
        paragraph.runs[0].font.size = Pt(14)
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    def set_rtl_paragraph(self, para) -> None:
        """Ensures the paragraph and its text are in RTL to fix punctuation issues."""
        pPr = para._element.get_or_add_pPr()

        # Set paragraph direction to RTL
        rtl_dir = OxmlElement("w:bidi")  # 'bidi' ensures bidirectional text handling
        rtl_dir.set(qn("w:val"), "1")
        pPr.append(rtl_dir)

        # Ensure all runs are also RTL
        for run in para.runs:
            rPr = run._element.get_or_add_rPr()
            rtl_run = OxmlElement("w:rtl")
            rtl_run.set(qn("w:val"), "1")
            rPr.append(rtl_run)

        # Set alignment to right
        alignment = OxmlElement("w:jc")
        alignment.set(qn("w:val"), "left")
        pPr.append(alignment)

    def replace_braces(self, val: str) -> str:
        """switches each brace to the opposite brace"""
        if isinstance(val, str):
            val = val.translate(
                str.maketrans(
                    {"(": ")", ")": "(", "{": "}", "}": "{", "[": "]", "]": "["}
                )
            )
        return val

    def add_literals(
            self, doc: DocxTemplate, literals, person_name, literal_ouptut_path
    ) -> None:
        """Add the literals to the document."""
        context = {
            "improve": {"points": literals[person_name]["points to improve"]},
            "conserve": {"points": literals[person_name]["points to conserve"]},
        }

        context["improve"]["points"] = [
            ensure_ends_with(point, ".")
            for point in context["improve"]["points"]
            if (isinstance(point, str) and len(point) > 4)
        ]
        context["conserve"]["points"] = [
            ensure_ends_with(point, ".")
            for point in context["conserve"]["points"]
            if (isinstance(point, str) and len(point) > 4)
        ]

        doc.render(context)
        doc.save(literal_ouptut_path)

    def hash_file(self, names_to_hashes, person_name, n=None) -> str:
        """Hash the file name and save it to a file if names_to_hashes is True. returns the title to save the file as."""
        if names_to_hashes:
            with open("names_to_hashes.txt", "a", encoding="utf-8") as f:
                f.write(
                    f"{person_name} ≥ {self.my_hash(person_name).encode('unicode_escape').decode('utf-8')}\n"
                )
            title_to_save = (
                f"{self.my_hash(person_name)} (N={n})".replace('"', "").replace("'", "")
                + ".docx"
            )
        else:
            title_to_save = (
                f"{person_name} (N={n})".replace('"', "").replace("'", "") + ".docx"
            )

        return title_to_save

    def create_word_file(
        self,
        averages,
        stds,
        counts,
        high_knowing,
        low_knowing,
        literals,
        person_name,
        n=None,
        names_to_hashes=False,
    ) -> None:
        """Create a Word file for a specific person with the given data."""

        # hash the file name and save it to a file if names_to_hashes is True
        title_to_save = self.hash_file(names_to_hashes, person_name, n)
        path_to_save = os.path.join(self.word_output_dir, title_to_save)

        TEMP_FILE_PATH = "tmp.docx"
        template = DocxTemplate(self.file_format_path)
        self.add_literals(template, literals, person_name, TEMP_FILE_PATH)

        doc = Document(TEMP_FILE_PATH)

        # title and font
        title_str = (
            "N="
            + str(n)
            + " ,"
            + "שיקוף סוציומטרי - "
            + person_name
            + ", סמסטר "
            + str(constants.SEMESTER)
        )
        self.set_title(doc, title_str)

        self.fill_main_table(doc, averages, stds, counts, person_name)
        self.fill_values_table(doc, counts, person_name)
        self.fill_semester_table(doc, averages, person_name) # MARK IMPORTANT

        # add the main graph to the table
        TMP_FILE_PATH = "tmp.png"
        self.create_main_graph(
            doc,
            averages,
            stds,
            person_name,
            "images/" + person_name + ".png",
            "main_graph_professional",
            categories=constants.PROFESSIONAL_CATEGORIES,
            title="מדדים ביצועיים",
        )

        # MARK IMPORTANT
        self.create_progress_graph(
            doc, averages, person_name, TMP_FILE_PATH, "progress_graph"
        )

        self.create_knowing_graph(
            doc, high_knowing, low_knowing, person_name, TMP_FILE_PATH, "knowing_graph"
        )

        # Save the document
        doc.save(path_to_save)
        os.remove(TEMP_FILE_PATH)

    def get_numerical_columns(self, df: pd.DataFrame) -> list:
        # Get the numerical columns from the dataframe
        numerical_columns = df.select_dtypes(include=[np.number]).columns.tolist()
        return numerical_columns

    def calculate_averages(self, combined_dfs):
        # Calculate stds and averages for each person in the combined_df
        last_combined_df = combined_dfs[-1]
        last_df_names = last_combined_df["name"].unique()

        averages = {}
        high_knowing = {}
        low_knowing = {}
        counts = {}
        counts["total"] = [
            {0: {}, 1: {}, 2: {}, 3: {}, 4: {}, 5: {}, 6: {}, "total": {}}
        ] * len(combined_dfs)
        stds = {}

        for idx, combined_df in enumerate(combined_dfs):

            # calculate total counts
            for i in range(7):
                for category in constants.CATEGORY_NAME_DICT.keys():
                    counts["total"][idx][i][category] = combined_df[
                        combined_df[category] == i
                    ].shape[0]

            for person_name, group in combined_df.groupby("name"):

                # if person_name == "נדב לוי":
                #     continue

                name_fix_dict = {
                    "גלעד הרצברג": "גלעד הרצברג רבינוביץ'",
                    "ליאם מור": "ליאם ארי מור",
                    "אייל שלמה אפרימה": "אייל אפרימה",
                    "אייל ווינטרויב": "אייל ויינטרוב",
                    "איתי אהרון פייביש": "איתי פייביש",
                    "אליאב מנחם אופטובסקי": "אליאב אופטובסקי",
                    "זהר רטנר שרף": "זהר רטנר",
                    "יואב סטרולוביץ": "יואב סטרולוביץ'",
                    "עידו דיוידסון רומנו": "עידו רומנו",
                    "תומר יוסף גרונר": "תומר גרונר",
                }

                if person_name in name_fix_dict:
                    person_name = name_fix_dict[person_name]
                    print("enters")

                if person_name not in last_df_names:
                    print(
                        f"Person {person_name} not found in the last combined dataframe."
                    )
                    continue

                # Initialize the dictionaries for each person if not already done
                if person_name not in averages:
                    averages[person_name] = []
                    high_knowing[person_name] = []
                    low_knowing[person_name] = []

                    if person_name != "total":
                        counts[person_name] = []

                    stds[person_name] = []

                # Append empty dictionaries for each index
                stds[person_name].append({})
                averages[person_name].append({})
                high_knowing[person_name].append({})
                low_knowing[person_name].append({})



                if person_name != "total":
                    counts[person_name].append({})



                    counts[person_name][idx] = {
                        0: {},
                        1: {},
                        2: {},
                        3: {},
                        4: {},
                        5: {},
                        6: {},
                        "total": {},
                    }

                numerical_columns = self.get_numerical_columns(combined_df)

                # calculating the averages and stds for each category
                for category in numerical_columns:
                    group_without_0 = group[
                        (group[category] >= 1) & (group[category] <= 6)
                    ]
                    # Calculate new stds
                    std = group_without_0[category].std()
                    stds[person_name][idx][category] = std

                    # Calculate new averages
                    avg = group_without_0[category].mean()
                    averages[person_name][idx][category] = avg

                    # Calculate high and low knowing averages
                    high_knowing[person_name][idx][category] = group_without_0[
                        group_without_0["knowing"] > 4
                    ][category].mean()
                    low_knowing[person_name][idx][category] = group_without_0[
                        group_without_0["knowing"] <= 4
                    ][category].mean()

                    if person_name != "total":
                        # calculate count
                        for i in range(1, 7):
                            counts[person_name][idx][i][category] = group_without_0[
                                group_without_0[category] == i
                            ].shape[0]

                        # calculate total count
                        counts[person_name][idx]["total"][category] = (
                            group_without_0.shape[0]
                        )

            # Add a "total" person with the average of everyone
            if "total" not in averages:
                averages["total"] = []
                stds["total"] = []
                averages["range"] = []
                stds["range"] = []

            averages["total"].append({})
            averages["total"][idx] = {}
            averages["range"].append({})
            averages["range"][idx] = {}

            stds["total"].append({})
            stds["total"][idx] = {}

            for category in numerical_columns:
                all_avgs = [
                    averages[name][idx][category]
                    for name in averages.keys()
                    if name not in ["total", "range"]
                ]
                min_total = min(all_avgs)
                max_total = max(all_avgs)
                print("the min and max are:", min_total, ", ", max(all_avgs))

                averages["total"][idx][category] = np.average(all_avgs)
                averages["range"][idx][category] = (min_total, max_total)

                stds["total"][idx][category] = np.std(all_avgs)

        return averages, stds, counts, high_knowing, low_knowing

    def get_literals(self, combined_df: pd.DataFrame) -> list:

        # get the literal columns from the dataframe
        literal_columns = combined_df.columns.drop(
            self.get_numerical_columns(combined_df)
        ).drop("name")
        literals = {}

        for person_name, group in combined_df.groupby("name"):
            # Initialize the dictionary for each person if not already done
            if person_name not in literals:
                literals[person_name] = {
                    "points to conserve": [],
                    "points to improve": [],
                }

            # adding both improve and conserve points
            for category in literal_columns:
                for index, value in group[category].dropna().items():
                    knowing_value = group.loc[index, "knowing"]
                    value = ensure_ends_with(value, ADD_IN_END_OF_SENTENCE)

                    # adding the corresponding finish based on the knowing value
                    if knowing_value > 4:
                        literals[person_name][category].append(
                            self.replace_braces(f"{value} (מידת היכרות גבוהה)")
                        )
                    else:
                        literals[person_name][category].append(
                            self.replace_braces(f"{value} (מידת היכרות נמוכה)")
                        )

        return literals

    def run_word_creation(
        self,
        combined_dfs: pd.DataFrame,
        name_to_classification: dict = None,
        verbose: bool = True,
        names_to_hashes: bool = False,
        is_commander_file: bool = True,
    ):

        averages, stds, counts, high_knowing, low_knowing = self.calculate_averages(
            combined_dfs
        )

        # generate commander file graphs
        if is_commander_file:
            commander_file_helper.generate_all_graphs(averages, stds, counts)

        else:

            literals = self.get_literals(combined_dfs[-1])
            # create word file for every person
            for df in tqdm(combined_dfs[-1].groupby("name"), disable=not verbose):
                person_name = df[0]

                # if person_name == "נדב לוי":
                #     continue

                self.create_main_graph(
                    doc=None,
                    averages=averages,
                    stds=stds,
                    person_name=person_name,
                    path_to_save=f"images/{person_name}_main.png",
                    tag=None,
                    categories=constants.PROFESSIONAL_CATEGORIES,
                    title="מדדים ביצועיים",
                )

                self.create_knowing_graph(
                    doc=None,
                    high_knowing=high_knowing,
                    low_knowing=low_knowing,
                    person_name=person_name,
                    path_to_save=f"images/{person_name}_knowing.png",
                    tag=None,
                )

                # continue

                self.create_word_file(
                    averages,
                    stds,
                    counts,
                    high_knowing,
                    low_knowing,
                    literals,
                    person_name,
                    n=df[1].shape[0],
                )



