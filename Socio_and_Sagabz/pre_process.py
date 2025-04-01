from abc import ABC, abstractmethod
import warnings
import pandas as pd
import os

class PreProcess(ABC):
    def __init__(self):
        self.column_names = self.get_column_names()
    
    
    def gen_combined_dataframe(self, input_dir_path):
        self.input_dir = input_dir_path

        # sanity checks
        if not os.path.isdir(self.input_dir):
            raise Exception(f"the path is not a dir: {self.input_dir}")

        # process all excel files as dataframes
        dataframes = []
        for file_path in os.listdir(self.input_dir):
            # weird thing that it showed semi open files (with the $ symbol - denoting they are open)
            if "$" in file_path:
                continue

            # check if the file is an excel file
            if not file_path.endswith(".xlsx"):
                warnings.warn(f"non excel file in the dir: {file_path}")
                warnings.warn("skipping the file")
                continue

            df = pd.read_excel(os.path.join(self.input_dir, file_path))
            df = self.remove_unwanted_data(df)
            df = df.rename(columns={df.columns[i]: self.column_names[i] for i in range(len(self.column_names))})

            dataframes.append(df)

        # combine everything to one single big table
        # Iterate through the remaining dataframes and update the base dataframe
        return self.combine_dataframes(dataframes)
    
    
    def gen_data_per_person(self, combined_dataframe):
        data_per_person = []

        for df in combined_dataframe.groupby(by="name"):
            data_per_person.append(df[1])
        
        return data_per_person
        
    

    def run(self, input_dir_path):
        self.input_dir = input_dir_path
        combined_dataframe = self.gen_combined_dataframe(input_dir_path)

        # drop all rows that the group by "name" in groupby is less than 3
        # if not testing:
        #     combined_dataframe = combined_dataframe.groupby("name").filter(lambda x: len(x) >= 3)
        
        self.gen_data_per_person(combined_dataframe)

        data_per_person = self.gen_data_per_person(combined_dataframe)
        
        return [combined_dataframe], data_per_person


    def remove_unwanted_data(self, df: pd.DataFrame):
        # check if there is a timestamp column, and remove it
        first_column_name = df.columns[0]
        if first_column_name.startswith("חותמת"):
            df.drop(columns=first_column_name, inplace=True)

        # now check for segel column and remove it
        first_column_name = df.columns[0]
        if first_column_name.startswith("אני מסגל"):
            df.drop(columns=first_column_name, inplace=True)

        try:
            cols_to_remove = [col for col in df.columns if ("Unnamed" in col) or ("Column" in col)]
        except TypeError:
            raise Exception("Type Error in the column names, probably someone put a number in one of the column names,"
                            "\n you should find out what excel it is and look at its columns")
        if len(cols_to_remove) != 0:
            df.drop(columns=cols_to_remove, inplace=True)

        # remove empty lines
        df = df.dropna(axis=0, thresh=2)

        df = self.remove_platform_specific(df)
        return df

    def combine_dataframes(self, dataframes):
        return pd.concat(dataframes, ignore_index=True, axis=0)

    @abstractmethod
    def remove_platform_specific(self, df: pd.DataFrame):
        pass

    @abstractmethod
    def get_column_names(self):
        pass
