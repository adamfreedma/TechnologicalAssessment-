import os
import pprint
from abc import ABC, abstractmethod
from tkinter import Tk
from tkinter.filedialog import askdirectory
import pandas as pd
import numpy as np
import tqdm

from statistics_socio import Statistics
from column_constants import SIGMAS
import socio_to_classification


class SocioAndSagabz:
    def __init__(
        self,
        inputs_path,
        outputs_path,
        format_path,
        testing=False,
        combine_excels=False,
        run_classification=True,
        names_to_hashes=False,
        is_commander_file=False,
    ):
        self.inputs_path = inputs_path
        self.outputs_path = outputs_path
        self.format_path = format_path
        self.testing = testing
        self.run_classification = run_classification

        self.combine_excels = combine_excels
        self.split_excels = True
        self.run_statistics = True
        self.is_commander_file = is_commander_file

        self.raw_data_dir_path = os.path.join(self.outputs_path, "raw_data")
        self.combined_df = []
        self.data_per_person_list = None
        self.name_to_classification = None
        self.stats_df = None
        self.names_to_hashes = names_to_hashes

    @abstractmethod
    def get_preprocess_obj(self):
        pass

    @abstractmethod
    def get_docx_obj(self):
        pass

    def get_file_format_path(self):
        return self.format_path

    def ensure_dir(self, path):
        if not os.path.isdir(path):
            os.mkdir(path)

    def export_to_excel(self, df, file_title):
        df.to_excel(file_title, index=False)

    def load_from_excel(self, path):
        return pd.read_excel(path, header=0, engine="openpyxl")

    def preprocess_and_combine(self):
        print("---------------------------------")
        print("Combining and Preprocessing Data")
        preprocess_obj = self.get_preprocess_obj()()
        combined_data_path = os.path.join(self.outputs_path, "combined_data.xlsx")

        if self.combine_excels:
            self.combined_df, self.data_per_person_list = preprocess_obj.run(
                self.inputs_path
            )
            self.export_to_excel(self.combined_df[-1], combined_data_path)

        else:
            for semester in range(1, 7):
                combined_data_path = os.path.join(
                    self.inputs_path, f"combined_data_{semester}.xlsx"
                )
                if os.path.exists(combined_data_path):
                    print(combined_data_path)
                    self.combined_df.append(self.load_from_excel(combined_data_path))
            self.data_per_person_list = preprocess_obj.gen_data_per_person(
                self.combined_df[-1]
            )

    def create_statistics(self):
        # create stats excel, and save
        print("---------------------------------")
        print("creating stats excel for everyone")
        stats_path = os.path.join(self.outputs_path, "stats_excel.xlsx")

        if self.run_statistics:
            stat_obj = Statistics()
            self.stats_df = stat_obj.run(self.combined_df[-1])

            self.stats_df["n_sigma"] = self.stats_df.groupby("category")[
                "mean"
            ].transform(lambda x: (x - np.mean(x)) / np.std(x))

            self.export_to_excel(self.stats_df, stats_path)
        else:
            self.stats_df = self.load_from_excel(stats_path)

    def create_word_files(self):

        # save word files
        word_output_dir = os.path.join(self.outputs_path, "word")
        self.ensure_dir(word_output_dir)

        print("making word files")
        # create word files
        docx_obj = self.get_docx_obj()(self.format_path, word_output_dir)
        
        docx_obj.run_word_creation(
            self.combined_df,
            name_to_classification=self.name_to_classification,
            names_to_hashes=self.names_to_hashes,
            is_commander_file=self.is_commander_file,
        )

    def save_sigmas(self):
        """
        Save the threshold for large (top 15%) and small (last 15%) sigma values for each category
        in the column_constants file SIGMAS constant.
        :return: None
        """
        stats_path = os.path.join(self.outputs_path, "stats_excel.xlsx")
        stats_df = self.load_from_excel(stats_path)

        sigmas_df = stats_df.groupby("category")["std"]
        # cumpute the 85% and 15% quantile for each category and save it in the SIGMAS constant
        for category, sigma in sigmas_df:
            SIGMAS[category] = (sigma.quantile(0.15), sigma.quantile(0.85))

    def run(self):
        self.ensure_dir(self.outputs_path)
        self.preprocess_and_combine()
        self.create_statistics()
        self.save_sigmas()

        if not self.combine_excels:
            self.create_word_files()
