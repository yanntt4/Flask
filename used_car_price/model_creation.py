# -*- coding: utf-8 -*-
"""
Created on Fri Nov 17 13:12:32 2023

@author: ythiriet
"""


# Global importation
import sys
import math
import matplotlib.pyplot as plot
import pandas as pd
import numpy as np
from tqdm import tqdm
from sklearn.model_selection import KFold
import random
from sklearn.preprocessing import OneHotEncoder, RobustScaler
import joblib
import numbers

from sklearn.impute import KNNImputer
from sklearn.neighbors import KNeighborsRegressor
from sklearn.metrics.pairwise import cosine_similarity

# Personal library
sys.path.append("./class")
from Data_modelling import Data_modelling
from Data_modelling_random_forest import random_forest
from Data_modelling_gradient_boosting import gradient_boosting
from Data_modelling_neural_network import neural_network
from Data_modelling_xgboosting import xgboosting
from Data_modelling_catboosting import catboosting
from Data_modelling_lightboosting import lightboosting
from Data_modelling_stacking import stacking



# Class containing all parameters
class Parameters():
    def __init__(self):
        self.CLEAR_MODE = True
        
        self.NAME_DATA_PREDICT = "price"
        self.GENERIC_NAME_DATA_PREDICT = "Car price" # for plot

        self.SWITCH_REMOVING_DATA = True
        self.LIST_DATA_DROP = ["id"]
        self.SWITCH_DATA_REDUCTION = False
        self.SWITCH_DATA_NOT_ENOUGHT = False
        self.NB_DATA_NOT_ENOUGHT = 1500
        self.SWITCH_ABERRANT_IDENTICAL_DATA = True
        self.SWITCH_RELATION_DATA = False
        self.ARRAY_RELATION_DATA = np.array([["Height", 2],["Age", 2]], dtype = object)
        
        self.SWITCH_DATA_SCALING = True
        self.LIST_DATA_SCALING = ["milage"]
        
        self.SWITCH_ENCODE_DATA_PREDICT = False
        self.ARRAY_DATA_ENCODE_PREDICT = np.array([[self.NAME_DATA_PREDICT,"p",1],[self.NAME_DATA_PREDICT,"e",0]], dtype = object)
        self.SWITCH_ENCODE_DATA = True
        self.LIST_DATA_ENCODE_REPLACEMENT = ["brand", "fuel_type", "transmission","ext_col","int_col", "accident", "clean_title", "engine", "model"]
        self.SWITCH_ENCODE_DATA_ONEHOT = False
        self.LIST_DATA_ENCODE_ONEHOT = ["Gender", "Vehicle_Age", "Vehicle_Damage"]
        self.SWITCH_ENCODE_DATA_MANUAL_REPLACEMENT = False
        self.ARRAY_DATA_ENCODE_REPLACEMENT = np.zeros(3, dtype = object)
        self.ARRAY_DATA_ENCODE_REPLACEMENT[0] = np.array(
            [["Gender","Female",1,"FEMME"],["Gender","Male",0,"HOMME"]], dtype = object)
        self.ARRAY_DATA_ENCODE_REPLACEMENT[1] = np.array(
            [["Vehicle_Age","< 1 Year",0,"Neuf"],["Vehicle_Age","1-2 Year",1,"Occasion/Neuf"],["Vehicle_Age","> 2 Years",1,"Occasion"]], dtype = object)
        self.ARRAY_DATA_ENCODE_REPLACEMENT[2] = np.array(
            [["Vehicle_Damage","Yes",1,"OUI"],["Vehicle_Damage","No",0,"NON"]], dtype = object)

        self.SWITCH_PLOT_DATA = True
        
        # Classification and Regression
        self.SWITCH_OVERSAMPLING = False
        # Regression
        self.SWITCH_SMOTER_DATA = False
        self.SMOTER_THRESHOLD = 100000
        # Classification
        self.SWITCH_SMOTEN_DATA = False
        
        self.SWITCH_REPLACING_NAN = False
        self.SWITCH_SAMPLE_DATA = False
        self.FRACTION_SAMPLE_DATA = 0.01

        self.MODEL_SAVING = False
    
        self.RF_MODEL = False
        self.RF_MODEL_OPTI = False
        self.RF_MODEL_TRIAL = 25
        
        self.GB_MODEL = False
        self.GB_MODEL_OPTI = False
        self.GB_MODEL_TRIAL = 25

        self.NN_MODEL = False
        self.NN_MODEL_OPTI = False
        self.NN_MODEL_TRIAL = 5

        self.XG_MODEL = False
        self.XG_MODEL_OPTI = False
        self.XG_MODEL_TRIAL = 200
        
        self.CB_MODEL = False
        self.CB_MODEL_OPTI = False
        self.CB_MODEL_TRIAL = 50
        
        self.LB_MODEL = False
        self.LB_MODEL_OPTI = False
        self.LB_MODEL_TRIAL = 250

        self.STACKING = False
        if self.STACKING:
            self.XG_MODEL = True
            self.CB_MODEL = True
            self.LB_MODEL = True

        self.MULTI_CLASSIFICATION = False

        self.N_SPLIT = 5
        self.k_folds = KFold(n_splits=self.N_SPLIT)


    # Determining if multi-classification
    def multi_classification_analysis(self, UNIQUE_PREDICT_VALUE):

        if UNIQUE_PREDICT_VALUE.shape[0] > 2:
            self.MULTI_CLASSIFICATION = True
    
    
    def regression_analysis(self, TRAIN_DATAFRAME):
        if isinstance(TRAIN_DATAFRAME[self.NAME_DATA_PREDICT][0], numbers.Number):
            self.REGRESSION = True
        else:
            self.REGRESSION = False
    
    def saving_array_replacement(self):
        joblib.dump(self.ARRAY_DATA_ENCODE_REPLACEMENT, "./data_replacement/array_data_encode_replacement.joblib")
        

class Data_Preparation():
    def __init__(self):
        self.TRAIN_DATAFRAME = []
        self.TEST_DATAFRAME = []
        self.TRAIN_STATS = []
        self.UNIQUE_PREDICT_VALUE = []
        self.TRAIN_CORRELATION = []
        self.DUPLICATE_LINE = []

        self.ARRAY_REPLACEMENT_ALL = np.zeros([0], dtype = object)
        self.INDEX_REPLACEMENT_ALL = np.zeros([0], dtype = object)


    def data_import(self, NAME_DATA_PREDICT):

        self.TRAIN_DATAFRAME = pd.read_csv("./data/train.csv")
        self.TEST_DATAFRAME = pd.read_csv("./data/test.csv")
        self.TRAIN_STATS = self.TRAIN_DATAFRAME.describe()


    def data_predict_description(self, NAME_DATA_PREDICT):
        self.UNIQUE_PREDICT_VALUE = self.TRAIN_DATAFRAME.groupby(NAME_DATA_PREDICT)[NAME_DATA_PREDICT].count()
        
        # Printing first values
        print(self.TRAIN_DATAFRAME.head())
    
    
    def data_robust_scaler(self, COLUMN_NAME):
        
        for COLUMN in COLUMN_NAME:
            transformer = RobustScaler().fit(pd.DataFrame(pd.concat([self.TRAIN_DATAFRAME[COLUMN], self.TEST_DATAFRAME[COLUMN]],ignore_index=True)))
            self.TRAIN_DATAFRAME[COLUMN] = transformer.transform(pd.DataFrame(self.TRAIN_DATAFRAME[COLUMN]))
            self.TEST_DATAFRAME[COLUMN] =transformer.transform(pd.DataFrame(self.TEST_DATAFRAME[COLUMN]))


    def data_encoding_replacement(self, ARRAY_REPLACEMENT, NAN_VALUES = False):
        
        for i_encoding, DataFrame in enumerate([self.TRAIN_DATAFRAME, self.TEST_DATAFRAME]):
    
            # Replacement
            for j in range(ARRAY_REPLACEMENT.shape[0]):
                for k in range(ARRAY_REPLACEMENT[j].shape[0]):
                    DataFrame[ARRAY_REPLACEMENT[j][k][0]] = DataFrame[ARRAY_REPLACEMENT[j][k][0]].replace(
                        ARRAY_REPLACEMENT[j][k][1], int(ARRAY_REPLACEMENT[j][k][2]))
                
            # Replacing nan values
            if NAN_VALUES:
                DataFrame[ARRAY_REPLACEMENT[j][0][0]] = DataFrame[ARRAY_REPLACEMENT[j][0][0]].fillna(0)

            # Recording the replacement
            if i_encoding == 0:
                self.TRAIN_DATAFRAME = DataFrame
            else:
                self.TEST_DATAFRAME = DataFrame


    def data_encoding_replacement_factorize(self, COLUMN_NAME):

        # Init
        ARRAY_DATA_ENCODE_REPLACEMENT = np.zeros([len(COLUMN_NAME)], dtype = object)
        
        for i, COLUMN in tqdm(enumerate(COLUMN_NAME)):
            
            # Tranforming data
            TRANSFORM_ARRAY, DATA_ENCODE_REPLACEMENT = pd.factorize(pd.concat([self.TRAIN_DATAFRAME[COLUMN], self.TEST_DATAFRAME[COLUMN]],
                                                                              ignore_index=True))
            self.TRAIN_DATAFRAME[COLUMN] = pd.DataFrame(TRANSFORM_ARRAY[:int(self.TRAIN_DATAFRAME.shape[0] + 1)], columns = [COLUMN])
                      
            # Saving transformation using specific format
            DATA_ENCODE_REPLACEMENT = pd.DataFrame(DATA_ENCODE_REPLACEMENT, columns = ["Encoding"])
            DATA_ENCODE_REPLACEMENT["Index"] = DATA_ENCODE_REPLACEMENT.index
            DATA_ENCODE_REPLACEMENT[COLUMN] = COLUMN
            ARRAY_DATA_ENCODE_REPLACEMENT[i] = np.array(DATA_ENCODE_REPLACEMENT)
            ARRAY_DATA_ENCODE_REPLACEMENT[i] = np.array(pd.DataFrame(ARRAY_DATA_ENCODE_REPLACEMENT[i]).iloc[:,[2,0,1]])
        
        # Exit
        return ARRAY_DATA_ENCODE_REPLACEMENT


    def data_encoding_replacement_predict(self, ARRAY_REPLACEMENT):
        for j in range(ARRAY_REPLACEMENT.shape[0]):
            self.TRAIN_DATAFRAME[ARRAY_REPLACEMENT[j,0]] = self.TRAIN_DATAFRAME[ARRAY_REPLACEMENT[j,0]].replace(
                ARRAY_REPLACEMENT[j,1],ARRAY_REPLACEMENT[j,2])


    def data_encoding_onehot(self, NAME_DATA_ENCODE):
        Enc = OneHotEncoder(handle_unknown='ignore')
        DATA_ENCODE_TRAIN = self.TRAIN_DATAFRAME.loc[:,[NAME_DATA_ENCODE]]
        DATA_ENCODE_TEST = self.TEST_DATAFRAME.loc[:,[NAME_DATA_ENCODE]]
        DATA_ENCODE_NAME = pd.Series(NAME_DATA_ENCODE + DATA_ENCODE_TRAIN.groupby(NAME_DATA_ENCODE)[NAME_DATA_ENCODE].count().index)
        DATA_ENCODE_NAME = DATA_ENCODE_NAME.replace(["<",">"]," ", regex=True)
        Enc.fit(DATA_ENCODE_TRAIN)

        DATA_ENCODE_TRAIN = Enc.transform(DATA_ENCODE_TRAIN).toarray()
        DATA_ENCODE_TRAIN = pd.DataFrame(DATA_ENCODE_TRAIN, columns = DATA_ENCODE_NAME)
        DATA_ENCODE_TRAIN = DATA_ENCODE_TRAIN.set_index(self.TRAIN_DATAFRAME.index)

        self.TRAIN_DATAFRAME = self.TRAIN_DATAFRAME.drop(columns = NAME_DATA_ENCODE)
        self.TRAIN_DATAFRAME = pd.concat([self.TRAIN_DATAFRAME, DATA_ENCODE_TRAIN], axis = 1)

        DATA_ENCODE_TEST = Enc.transform(DATA_ENCODE_TEST).toarray()
        DATA_ENCODE_TEST = pd.DataFrame(DATA_ENCODE_TEST, columns = DATA_ENCODE_NAME)
        DATA_ENCODE_TEST = DATA_ENCODE_TEST.set_index(self.TEST_DATAFRAME.index)

        self.TEST_DATAFRAME = self.TEST_DATAFRAME.drop(columns = NAME_DATA_ENCODE)
        self.TEST_DATAFRAME = pd.concat([self.TEST_DATAFRAME, DATA_ENCODE_TEST], axis = 1)
    
    
    def encode_data_error_removal(self, ARRAY_REPLACEMENT):
        for ARRAY in ARRAY_REPLACEMENT:
            Global_Data.TRAIN_DATAFRAME[ARRAY[0][0]] = pd.to_numeric(
                Global_Data.TRAIN_DATAFRAME[ARRAY[0][0]],errors="coerce", downcast = 'integer')
    
    
    def data_format_removal(self, ARRAY_REPLACEMENT, Type = str, Len = 2):
        for ARRAY in ARRAY_REPLACEMENT:
            Global_Data.TRAIN_DATAFRAME = Global_Data.TRAIN_DATAFRAME.loc[(
                Global_Data.TRAIN_DATAFRAME[ARRAY[0][0]].astype(Type).str.len() < Len)]
        

    def data_drop(self, Name_data_drop):
        self.TRAIN_DATAFRAME = self.TRAIN_DATAFRAME.drop([Name_data_drop],axis=1)
        self.TEST_DATAFRAME = self.TEST_DATAFRAME.drop([Name_data_drop],axis=1)


    def data_pow(self, Name_Data_Duplicate, Number_Duplicate):
        self.TRAIN_DATAFRAME[Name_Data_Duplicate] = (
            self.TRAIN_DATAFRAME[Name_Data_Duplicate].pow(Number_Duplicate))
        self.TEST_DATAFRAME[Name_Data_Duplicate] = (
            self.TEST_DATAFRAME[Name_Data_Duplicate].pow(Number_Duplicate))


    def data_duplicate_removal(self, NAME_DATA_PREDICT, Column_Drop = ""):

        if len(Column_Drop) == 0:
            Duplicated_Data_All = self.TRAIN_DATAFRAME.drop(NAME_DATA_PREDICT, axis = 1).duplicated()
        else:
            Duplicated_Data_All = self.TRAIN_DATAFRAME.drop([Column_Drop, NAME_DATA_PREDICT],axis = 1).duplicated()
        self.DUPLICATE_LINE = Duplicated_Data_All.loc[Duplicated_Data_All == True]
        self.TRAIN_DATAFRAME = self.TRAIN_DATAFRAME.drop(self.DUPLICATE_LINE.index)
        
        # Information to the user
        print(f"{self.DUPLICATE_LINE.shape[0]} has been removed because of duplicates")
        plot.pause(3)


    def remove_low_data(self, NB_DATA_NOT_ENOUGHT, NAME_DATA_NOT_ENOUGHT, LIST_NAME_DATA_REMOVE_MULTIPLE = []):

        # Searching for data with low values
        TRAIN_GROUP_VALUE = self.TRAIN_DATAFRAME.groupby(NAME_DATA_NOT_ENOUGHT)[NAME_DATA_NOT_ENOUGHT].count().sort_values(ascending = False)

        # Adding values only inside NAME DATA REMOVE MULTIPLE
        for NAME_DATA_REMOVE_MULTIPLE in LIST_NAME_DATA_REMOVE_MULTIPLE:
            TRAIN_GROUP_VALUE_OTHER = self.TRAIN_DATAFRAME.groupby(NAME_DATA_REMOVE_MULTIPLE)[NAME_DATA_REMOVE_MULTIPLE].count().index

        for VALUE_OTHER in TRAIN_GROUP_VALUE_OTHER:
            if np.sum(VALUE_OTHER == np.array(TRAIN_GROUP_VALUE.index)) == 0:
                TRAIN_GROUP_VALUE = pd.concat([TRAIN_GROUP_VALUE, pd.Series(0, index = [VALUE_OTHER])])

        # Searching for values to drop following number of elements
        REMOVE_TRAIN_GROUP_VALUE = TRAIN_GROUP_VALUE.drop(TRAIN_GROUP_VALUE[TRAIN_GROUP_VALUE > NB_DATA_NOT_ENOUGHT].index)

        # Removing data inside train and test dataframe
        for DATA_REMOVE in REMOVE_TRAIN_GROUP_VALUE.index:
            self.TRAIN_DATAFRAME = self.TRAIN_DATAFRAME.drop(self.TRAIN_DATAFRAME[self.TRAIN_DATAFRAME[NAME_DATA_NOT_ENOUGHT] == DATA_REMOVE].index)
            self.TEST_DATAFRAME = self.TEST_DATAFRAME.drop(self.TEST_DATAFRAME[self.TEST_DATAFRAME[NAME_DATA_NOT_ENOUGHT] == DATA_REMOVE].index)

            for NAME_DATA_REMOVE_MULTIPLE in LIST_NAME_DATA_REMOVE_MULTIPLE:
                self.TRAIN_DATAFRAME = self.TRAIN_DATAFRAME.drop(self.TRAIN_DATAFRAME[self.TRAIN_DATAFRAME[NAME_DATA_REMOVE_MULTIPLE] == DATA_REMOVE].index)
                self.TEST_DATAFRAME = self.TEST_DATAFRAME.drop(self.TEST_DATAFRAME[self.TEST_DATAFRAME[NAME_DATA_REMOVE_MULTIPLE] == DATA_REMOVE].index)

        # Reseting index
        self.TRAIN_DATAFRAME = self.TRAIN_DATAFRAME.reset_index(drop = True)
        self.TEST_DATAFRAME = self.TEST_DATAFRAME.reset_index(drop = True)


    def oversampling(self, NAME_DATA_PREDICT, NB_DATA_NOT_ENOUGHT, Name_Data_Oversample = ""):

        self.UNIQUE_PREDICT_VALUE = self.TRAIN_DATAFRAME.groupby(NAME_DATA_PREDICT)[NAME_DATA_PREDICT].count()
        Max_Nb_Data = np.amax(self.UNIQUE_PREDICT_VALUE.to_numpy())

        if len(Name_Data_Oversample) > 1:
            Global_Table_Train_Equilibrate = self.UNIQUE_PREDICT_VALUE.loc[(
                self.UNIQUE_PREDICT_VALUE.index == "Overweight_Level_II")]
        else:
            # Global_Table_Train_Equilibrate = self.UNIQUE_PREDICT_VALUE.loc[
            #     self.UNIQUE_PREDICT_VALUE > NB_DATA_NOT_ENOUGHT]
            Global_Table_Train_Equilibrate = self.UNIQUE_PREDICT_VALUE.loc[self.UNIQUE_PREDICT_VALUE < Max_Nb_Data]

        for i in tqdm(range(Global_Table_Train_Equilibrate.shape[0])):
            Matrix_To_Add = np.zeros(
                [0, self.TRAIN_DATAFRAME.shape[1]],
                dtype=object)
            DF_Reference = self.TRAIN_DATAFRAME.loc[self.TRAIN_DATAFRAME[NAME_DATA_PREDICT] == pd.DataFrame(
                Global_Table_Train_Equilibrate.index).iloc[i][0]]
            for j in tqdm(range(Max_Nb_Data - Global_Table_Train_Equilibrate.iloc[i])):
                Matrix_To_Add = np.append(
                    Matrix_To_Add,
                    np.zeros([1, self.TRAIN_DATAFRAME.shape[1]],
                              dtype=object),
                    axis=0)

                Matrix_To_Add[-1, :] = DF_Reference.iloc[
                    random.randint(0, DF_Reference.shape[0] - 1), :].to_numpy()

            DataFrame_To_Add = pd.DataFrame(
                Matrix_To_Add,
                columns=self.TRAIN_DATAFRAME.columns)

            self.TRAIN_DATAFRAME = pd.concat(
                [self.TRAIN_DATAFRAME, DataFrame_To_Add],
                ignore_index=True)


    def smoter(self, NAME_DATA_PREDICT, THRESHOLD):

        # implement SMOTER
        # see paper: https://core.ac.uk/download/pdf/29202178.pdf
        
        # Init
        DTYPE_DATAFRAME = self.TRAIN_DATAFRAME.dtypes


        # Sigmoid function
        def sigmoid(x):
          return 1 / (1 + np.exp(-x))


        # Applied sigmoid function to dataframe
        def relevance(x, threshold):
            x = np.array(x, dtype = float)
            y = x - threshold
            return sigmoid(y)


        # Create synthetic cases following closest neightbours
        def get_synth_cases(D, target, o=200, k=3, categorical_col = []):
            '''
            Function to generate the new cases.
            INPUT:
                D - pd.DataFrame with the initial data
                target - string name of the target column in the dataset
                o - oversampling rate
                k - number of nearest neighbors to use for the generation
                categorical_col - list of categorical column names
            OUTPUT:
                new_cases - pd.DataFrame containing new generated cases
            '''
            new_cases = pd.DataFrame(columns = D.columns) # initialize the list of new cases 
            ng = o // 100 # the number of new cases to generate
            for index, Case in tqdm(D.iterrows(), total=D.shape[0]):
                # find k nearest neighbors of the case
                knn = KNeighborsRegressor(n_neighbors = k+1) # k+1 because the case is the nearest neighbor to itself
                knn.fit(D.drop(columns = [target]).values, D[[target]])
                neighbors = knn.kneighbors(Case.drop(labels = [target]).values.reshape(1, -1), return_distance=False).reshape(-1)
                neighbors = np.delete(neighbors, np.where(neighbors == index))
                for i in range(0, ng):
                    # randomly choose one of the neighbors
                    x = D.iloc[neighbors[np.random.randint(k)]]
                    attr = np.zeros([1,11], dtype = object)          
                    for j, a in enumerate(D.columns):
                        # skip target column
                        if a == target:
                            continue;
                        if a in categorical_col:
                            # if categorical then choose randomly one of values
                            if np.random.randint(2) == 0:
                                attr[0,j] = Case[a]
                            else:
                                attr[0,j] = x[a]
                        else:
                            # if continious column
                            diff = Case[a] - x[a]
                            attr[0,j] = Case[a] + np.random.randint(2) * diff
                    # decide the target column
                    
                    
                    new = attr
                    d1 = cosine_similarity(new.reshape(1, -1), Case.drop(labels = [target]).values.reshape(1, -1))[0][0]
                    d2 = cosine_similarity(new.reshape(1, -1), x.drop(labels = [target]).values.reshape(1, -1))[0][0]
                    attr = np.append(attr, np.zeros([1,1], dtype = object), axis = 1)
                    attr[0,-1] = (d2 * Case[target] + d1 * x[target]) / (d1 + d2)
                    
                    
                    if i == 0:
                        new_cases = pd.DataFrame(attr, columns = D.columns)

                    
                    else:
                        new_cases = pd.concat([new_cases,pd.DataFrame(attr, columns = D.columns)], ignore_index=True)
                    
            # Exit
            return new_cases

        # Applied Smoter
        def SmoteR(D, target, threshold = 100000, 
                   th = 0.999, o = 200, u = 100, k = 3, categorical_col = []):
            '''
            The implementation of SmoteR algorithm:
            https://core.ac.uk/download/pdf/29202178.pdf
            INPUT:
                D - pd.DataFrame - the initial dataset
                target - the name of the target column in the dataset
                th - relevance threshold
                o - oversampling rate
                u - undersampling rate
                k - the number of nearest neighbors
            OUTPUT:
                new_D - the resulting new dataset
            '''
            # median of the target variable
            y_bar = D[target].median()
            
            # find rare cases where target less than median
            rareL = D[(relevance(D[target],threshold) > th) & (D[target] > y_bar)]  
            
            # generate rare cases for rareL
            new_casesL = get_synth_cases(rareL, target, o, k , categorical_col)
            
            # find rare cases where target greater than median
            rareH = D[(relevance(D[target],threshold) > th) & (D[target] < y_bar)]
            # generate rare cases for rareH
            new_casesH = get_synth_cases(rareH, target, o, k , categorical_col)
            
            new_cases = pd.concat([new_casesL, new_casesH], axis=0)
            
            # undersample norm cases
            norm_cases = D[relevance(D[target],threshold) <= th]
            # get the number of norm cases
            nr_norm = int(len(norm_cases) * u / 100)
            
            norm_cases = norm_cases.sample(min(len(D[relevance(D[target],threshold) <= th]), nr_norm))
            
            # get the resulting dataset
            new_D = pd.concat([new_cases, norm_cases], axis=0)
            
            # Exit
            return new_D


        # Smoter data
        np.random.seed(42)
        self.TRAIN_DATAFRAME = SmoteR(self.TRAIN_DATAFRAME, target=NAME_DATA_PREDICT, 
                                      threshold = THRESHOLD,
                                      th = 0.999, o = 300, u = 100, k = 10)
        
        # Changing column type after smoter
        for COLUMN,TYPE in zip(DTYPE_DATAFRAME.index, DTYPE_DATAFRAME):
            self.TRAIN_DATAFRAME[COLUMN] = self.TRAIN_DATAFRAME[COLUMN].astype(TYPE)
        

    def data_sample(self, SAMPLE_FRACTION):

        self.TRAIN_DATAFRAME = self.TRAIN_DATAFRAME.sample(
            frac = SAMPLE_FRACTION, replace = False, random_state = 42)


    def nan_replacing(self, COLUMN_NAMES):
        

        # Creating a column indicating missing value
        for COLUMN in COLUMN_NAMES:
            self.TRAIN_DATAFRAME[f"{COLUMN} Missing"] = self.TRAIN_DATAFRAME[f"{COLUMN}"].isnull()
        
        # Replacing missing values with nearest neightboor
        COLUMNS = self.TRAIN_DATAFRAME.columns
        imputer = KNNImputer(n_neighbors=20, weights="uniform")
        TRAIN_ARRAY = imputer.fit_transform(self.TRAIN_DATAFRAME)
        
        # Turning numpy array into dataframe
        self.TRAIN_DATAFRAME = pd.DataFrame(TRAIN_ARRAY, columns = COLUMNS)
    
    
    def saving_data_names(self):
        joblib.dump(self.TEST_DATAFRAME.columns, "./data_replacement/data_names.joblib")
        

class Data_Plot():
    def __init__(self):
        self.BOX_PLOT_DATA_PREDICT = ""
        self.BOX_PLOT_DATA_AVAILABLE = ""
        self.CORRELATION_PLOT = ""
        self.TRAIN_DATAFRAME = []
        self.TRAIN_CORRELATION = []
        self.UNIQUE_PREDICT_VALUE = []


    def box_plot_data_predict_plot(self, GENERIC_NAME_DATA_PREDICT, REGRESSION, NAME_DATA_PREDICT):
        
        if REGRESSION == False:

            # Init
            fig, self.BOX_PLOT_DATA_PREDICT = plot.subplots(2)
            plot.suptitle(f"Data count following {GENERIC_NAME_DATA_PREDICT}",
                          fontsize = 25,
                          color = "gold",
                          fontweight = "bold")

            # Horizontal bars for each possibilities
            self.BOX_PLOT_DATA_PREDICT[0].barh(
                y = self.UNIQUE_PREDICT_VALUE.index,
                width=self.UNIQUE_PREDICT_VALUE,
                height=0.03,
                label=self.UNIQUE_PREDICT_VALUE.index)
            
            # Cumulative horizontal bars
            Cumulative_Value = 0
            for i in range(self.UNIQUE_PREDICT_VALUE.shape[0]):
                self.BOX_PLOT_DATA_PREDICT[1].barh(
                    y=1,
                    width=self.UNIQUE_PREDICT_VALUE.iloc[i],
                    left = Cumulative_Value)
                self.BOX_PLOT_DATA_PREDICT[1].text(
                    x = Cumulative_Value + 100,
                    y = 0.25,
                    s = self.UNIQUE_PREDICT_VALUE.index[i])
                Cumulative_Value += self.UNIQUE_PREDICT_VALUE.iloc[i]
            self.BOX_PLOT_DATA_PREDICT[1].set_ylim(0, 2)
            self.BOX_PLOT_DATA_PREDICT[1].legend(
                self.UNIQUE_PREDICT_VALUE.index.to_numpy(),
                ncol=int(self.UNIQUE_PREDICT_VALUE.shape[0]/2),
                fontsize=6)
        
        # Regression
        else:
            # Init
            fig, self.BOX_PLOT_DATA_PREDICT = plot.subplots(2)
            plot.suptitle(f"Repartition of {GENERIC_NAME_DATA_PREDICT}",
                          fontsize = 25, color = "gold", fontweight = "bold")
            
            self.BOX_PLOT_DATA_PREDICT[0].hist(
                self.TRAIN_DATAFRAME[NAME_DATA_PREDICT], bins = int(self.TRAIN_DATAFRAME[NAME_DATA_PREDICT].shape[0]/20))
            self.BOX_PLOT_DATA_PREDICT[1].boxplot(self.TRAIN_DATAFRAME[NAME_DATA_PREDICT], 0, 'rs', 0)


    def plot_data_repartition(self):

        # Init
        NB_LINE = math.ceil((self.TRAIN_DATAFRAME.shape[1] - 1)/3)
        NB_COLUMN = math.ceil((self.TRAIN_DATAFRAME.shape[1] - 1)/3)

        # Box Plot for all data
        fig, self.BOX_PLOT_DATA_AVAILABLE = plot.subplots(NB_LINE, NB_COLUMN)
        plot.suptitle("Boxplot for all data into the TRAIN dataset",
                      fontsize = 25,
                      color = "chartreuse",
                      fontweight = "bold")

        for i in range(NB_LINE):
            for j in range(NB_COLUMN):
                if (i*NB_COLUMN + j < self.TRAIN_DATAFRAME.shape[1]) :
                    if (pd.unique(self.TRAIN_DATAFRAME.iloc[:, i*NB_COLUMN + j]).shape[0] < 400 or
                        (isinstance(self.TRAIN_DATAFRAME.iloc[:, i*NB_COLUMN + j][0], float) or
                        isinstance(self.TRAIN_DATAFRAME.iloc[:, i*NB_COLUMN + j][0], int))):
                        try:
                            self.BOX_PLOT_DATA_AVAILABLE[i, j].boxplot(
                                self.TRAIN_DATAFRAME.iloc[:, [i*NB_COLUMN + j]])
                            self.BOX_PLOT_DATA_AVAILABLE[i, j].set_title(
                                self.TRAIN_DATAFRAME.iloc[:, i*NB_COLUMN + j].name,
                                fontweight = "bold",
                                fontsize = 15)
                        except:
                            continue
    
    
    def plot_data_hist(self):

        # Init
        NB_LINE = math.ceil((self.TRAIN_DATAFRAME.shape[1] - 1)/3)
        NB_COLUMN = math.ceil((self.TRAIN_DATAFRAME.shape[1] - 1)/3)

        # Box Plot for all data
        fig, self.BOX_PLOT_DATA_AVAILABLE = plot.subplots(NB_LINE, NB_COLUMN)
        plot.suptitle("Boxplot for all data into the TRAIN dataset",
                      fontsize = 25,
                      color = "chartreuse",
                      fontweight = "bold")

        for i in range(NB_LINE):
            for j in range(NB_COLUMN):
                if (i*NB_COLUMN + j < self.TRAIN_DATAFRAME.shape[1]):
                    if (pd.unique(self.TRAIN_DATAFRAME.iloc[:, i*NB_COLUMN + j]).shape[0] < 400 or
                        (isinstance(self.TRAIN_DATAFRAME.iloc[:, i*NB_COLUMN + j][0], float)or
                        isinstance(self.TRAIN_DATAFRAME.iloc[:, i*NB_COLUMN + j][0], int))):
                        try:
                            self.BOX_PLOT_DATA_AVAILABLE[i, j].hist(
                                self.TRAIN_DATAFRAME.iloc[:, [i*NB_COLUMN + j]],
                                bins = 100)
                            self.BOX_PLOT_DATA_AVAILABLE[i, j].set_title(
                                self.TRAIN_DATAFRAME.iloc[:, i*NB_COLUMN + j].name,
                                fontweight = "bold",
                                fontsize = 15)
                        except:
                            continue


    def plot_data_relation(self, NAME_DATA_X, NAME_DATA_Y):

        plot.figure()
        plot.scatter(self.TRAIN_DATAFRAME[NAME_DATA_X], self.TRAIN_DATAFRAME[NAME_DATA_Y])
        plot.suptitle(
            f"Relation between {NAME_DATA_X} and {NAME_DATA_Y} variables",
            fontsize = 25,
            color = "darkorchid",
            fontweight = "bold")


    def CORRELATION_PLOT_Plot(self):

        fig2, self.CORRELATION_PLOT = plot.subplots()
        im = self.CORRELATION_PLOT.imshow(
            self.TRAIN_CORRELATION,
            vmin=-1,
            vmax=1,
            cmap="bwr")
        self.CORRELATION_PLOT.figure.colorbar(im, ax=self.CORRELATION_PLOT)
        self.CORRELATION_PLOT.set_xticks(np.linspace(
            0, self.TRAIN_DATAFRAME.shape[1] - 1, self.TRAIN_DATAFRAME.shape[1]))
        self.CORRELATION_PLOT.set_xticklabels(np.array(self.TRAIN_DATAFRAME.columns, dtype = str),
                                              rotation = 45)
        self.CORRELATION_PLOT.set_yticks(np.linspace(
            0, self.TRAIN_DATAFRAME.shape[1] - 1, self.TRAIN_DATAFRAME.shape[1]))
        self.CORRELATION_PLOT.set_yticklabels(np.array(self.TRAIN_DATAFRAME.columns, dtype = str))



# -- ////////// -- #
# -- ////////// -- #
# -- ////////// -- #





# Init for global parameters
Global_Parameters = Parameters()

if Global_Parameters.CLEAR_MODE:

    # Removing data
    from IPython import get_ipython
    get_ipython().magic('reset -sf')

    # Closing all figures
    plot.close("all")


Global_Data = Data_Preparation()
Global_Data.data_import(Global_Parameters.NAME_DATA_PREDICT)
Global_Parameters.regression_analysis(Global_Data.TRAIN_DATAFRAME)

HIGH_PRICE = 500000
print(f"Removing {Global_Data.TRAIN_DATAFRAME.loc[Global_Data.TRAIN_DATAFRAME['price'] > HIGH_PRICE].shape[0]} records")
Global_Data.TRAIN_DATAFRAME = Global_Data.TRAIN_DATAFRAME.loc[Global_Data.TRAIN_DATAFRAME["price"] < HIGH_PRICE].reset_index(drop = True)
Global_Data.TRAIN_DATAFRAME["accident"] = Global_Data.TRAIN_DATAFRAME["accident"].fillna("None reported")
Global_Data.TRAIN_DATAFRAME["clean_title"] = Global_Data.TRAIN_DATAFRAME["clean_title"].fillna("No")


# Droping some columns
if Global_Parameters.SWITCH_REMOVING_DATA:
    for name_drop in Global_Parameters.LIST_DATA_DROP:
        Global_Data.data_drop(name_drop)


# Removing variable with too low data
if Global_Parameters.SWITCH_DATA_REDUCTION:
    Global_Data.remove_low_data(
        Global_Parameters.NB_DATA_NOT_ENOUGHT, "Origin",
        LIST_NAME_DATA_REMOVE_MULTIPLE = ["Dest"])

# Data description
Global_Data.data_predict_description(Global_Parameters.NAME_DATA_PREDICT)

# Multi classification identification
Global_Parameters.multi_classification_analysis(Global_Data.UNIQUE_PREDICT_VALUE)


# Sample Data
if Global_Parameters.SWITCH_SAMPLE_DATA:
    Global_Data.data_sample(Global_Parameters.FRACTION_SAMPLE_DATA)


# Scaling Data
if Global_Parameters.SWITCH_DATA_SCALING:
    Global_Data.data_robust_scaler(Global_Parameters.LIST_DATA_SCALING)


# Encoding data for entry variables
if Global_Parameters.SWITCH_ENCODE_DATA:
    if Global_Parameters.SWITCH_ENCODE_DATA_ONEHOT:
        for NAME_DATA_ENCODE in Global_Parameters.LIST_DATA_ENCODE_ONEHOT:
            Global_Data.data_encoding_onehot(NAME_DATA_ENCODE)

    elif Global_Parameters.SWITCH_ENCODE_DATA_MANUAL_REPLACEMENT:
        # # Removing data with incorrect format before encoding
        # Global_Data.data_format_removal(Global_Parameters.ARRAY_DATA_ENCODE_REPLACEMENT)
        
        # Encoding
        Global_Data.data_encoding_replacement(Global_Parameters.ARRAY_DATA_ENCODE_REPLACEMENT, True)
    
        # Removing error data after encoding
        Global_Data.encode_data_error_removal(Global_Parameters.ARRAY_DATA_ENCODE_REPLACEMENT)
        Global_Data.TRAIN_DATAFRAME = Global_Data.TRAIN_DATAFRAME.dropna()
    
    else:
        
        Global_Parameters.ARRAY_DATA_ENCODE_REPLACEMENT = Global_Data.data_encoding_replacement_factorize(
            Global_Parameters.LIST_DATA_ENCODE_REPLACEMENT)


# Encoding data for predict variable
if Global_Parameters.SWITCH_ENCODE_DATA_PREDICT:
    Global_Data.data_encoding_replacement_predict(Global_Parameters.ARRAY_DATA_ENCODE_PREDICT)


# Searching for and removing aberrant/identical values
if Global_Parameters.SWITCH_ABERRANT_IDENTICAL_DATA:
    Global_Data.data_duplicate_removal(Global_Parameters.NAME_DATA_PREDICT)


# Oversampling to equilibrate data
if (Global_Parameters.SWITCH_OVERSAMPLING and
    Global_Parameters.SWITCH_SMOTEN_DATA == False and
    Global_Parameters.SWITCH_SMOTER_DATA == False):
    Global_Data.oversampling(Global_Parameters.NAME_DATA_PREDICT, Global_Parameters.NB_DATA_NOT_ENOUGHT)
elif (Global_Parameters.SWITCH_OVERSAMPLING == False and
    Global_Parameters.SWITCH_SMOTEN_DATA == False and
    Global_Parameters.SWITCH_SMOTER_DATA == True):
    Global_Data.smoter(Global_Parameters.NAME_DATA_PREDICT, Global_Parameters.SMOTER_THRESHOLD)


# Searching for repartition on data to predict
if Global_Parameters.SWITCH_PLOT_DATA:

    Global_Data_Plot = Data_Plot()
    Global_Data_Plot.TRAIN_DATAFRAME = Global_Data.TRAIN_DATAFRAME
    Global_Data_Plot.UNIQUE_PREDICT_VALUE = Global_Data.UNIQUE_PREDICT_VALUE
    Global_Data_Plot.box_plot_data_predict_plot(Global_Parameters.GENERIC_NAME_DATA_PREDICT,
                                                Global_Parameters.REGRESSION,
                                                Global_Parameters.NAME_DATA_PREDICT)
    Global_Data_Plot.plot_data_repartition()
    Global_Data_Plot.plot_data_hist()
    plot.pause(1)
    # Global_Data_Plot.plot_data_relation("Height", "Gender")
    plot.pause(1)
    Global_Data.TRAIN_CORRELATION = Global_Data.TRAIN_DATAFRAME.iloc[
        :,:Global_Data.TRAIN_DATAFRAME.shape[1] - 1].corr()
    Global_Data_Plot.TRAIN_CORRELATION = Global_Data.TRAIN_CORRELATION
    Global_Data_Plot.CORRELATION_PLOT_Plot()


# Modifying linear relation between data
if Global_Parameters.SWITCH_RELATION_DATA:
    for i in range(Global_Parameters.List_Relation_Data.shape[0]):
        Global_Data.data_pow(Global_Parameters.List_Relation_Data[i,0],
                             Global_Parameters.List_Relation_Data[i,1])


# Replacing Nan values
if Global_Parameters.SWITCH_REPLACING_NAN:
    Global_Data.nan_replacing(["BsmtQual", "BsmtCond", "BsmtExposure"])


# Generic Data Model
Data_Model = Data_modelling()
Data_Model.splitting_data(Global_Data.TRAIN_DATAFRAME,
                          Global_Parameters.NAME_DATA_PREDICT,
                          Global_Parameters.MULTI_CLASSIFICATION,
                          Global_Parameters.REGRESSION)
if (Global_Parameters.SWITCH_OVERSAMPLING == False and
    Global_Parameters.SWITCH_SMOTEN_DATA and
    Global_Parameters.SWITCH_SMOTER_DATA == False):
    Data_Model.smoten_sampling()


#
# Random Forest
if Global_Parameters.RF_MODEL:
    DATA_MODEL_RF = random_forest(Data_Model, Global_Parameters, Global_Data)


#
# Gradient Boosting
if Global_Parameters.GB_MODEL:
    DATA_MODEL_GB = gradient_boosting(Data_Model, Global_Parameters, Global_Data)
    

#
# Neural Network
if Global_Parameters.NN_MODEL:
    DATA_MODEL_NN = neural_network(Data_Model, Global_Parameters, Global_Data)


#
# XGBoosting
if Global_Parameters.XG_MODEL:
    DATA_MODEL_XG = xgboosting(Data_Model, Global_Parameters, Global_Data)


#
# Catboosting
if Global_Parameters.CB_MODEL:
    DATA_MODEL_CB = catboosting(Data_Model, Global_Parameters, Global_Data)


#
# Lightboosting
if Global_Parameters.LB_MODEL:
    DATA_MODEL_LB = lightboosting(Data_Model, Global_Parameters, Global_Data)



# Saving model and information
if Global_Parameters.MODEL_SAVING:
    Global_Parameters.saving_array_replacement()
    Global_Data.saving_data_names()


    if Global_Parameters.RF_MODEL:
        with open('./models/rf_model.sav', 'wb') as f:
            joblib.dump(DATA_MODEL_RF.MODEL, f)
    elif Global_Parameters.NN_MODEL:
        with open('./models/nn_model.sav', 'wb') as f:
            joblib.dump(DATA_MODEL_NN.MODEL, f)
    elif Global_Parameters.GB_MODEL:
        with open('./models/gb_model.sav', 'wb') as f:
            joblib.dump(DATA_MODEL_GB.MODEL, f)
    elif Global_Parameters.XG_MODEL:
        with open('./models/xg_model.sav', 'wb') as f:
            joblib.dump(DATA_MODEL_XG.MODEL, f)
    elif Global_Parameters.CB_MODEL:
        with open('./models/cb_model.sav', 'wb') as f:
            joblib.dump(DATA_MODEL_CB.MODEL, f)
    elif Global_Parameters.LB_MODEL:
        with open('./models/lb_model.sav', 'wb') as f:
            joblib.dump(DATA_MODEL_LB.MODEL, f)


# Kaggle competition preparation
Global_Data.TEST_DATAFRAME["accident"] = Global_Data.TEST_DATAFRAME["accident"].fillna("None reported")
Global_Data.TEST_DATAFRAME["clean_title"] = Global_Data.TEST_DATAFRAME["clean_title"].fillna("No")
for ARRAY in Global_Parameters.ARRAY_DATA_ENCODE_REPLACEMENT:
    for i in range(ARRAY.shape[0]):
        Global_Data.TEST_DATAFRAME[ARRAY[0,0]] = Global_Data.TEST_DATAFRAME[ARRAY[0,0]].replace(ARRAY[i,1],ARRAY[i,2])

# One model search
if Global_Parameters.STACKING == False:
    if Global_Parameters.RF_MODEL:
        KAGGLE_PREDICTION = pd.DataFrame(DATA_MODEL_RF.MODEL.predict(Global_Data.TEST_DATAFRAME), columns = [Global_Parameters.NAME_DATA_PREDICT])
    elif Global_Parameters.NN_MODEL:
        KAGGLE_PREDICTION = pd.DataFrame(DATA_MODEL_NN.MODEL.predict(Global_Data.TEST_DATAFRAME), columns = [Global_Parameters.NAME_DATA_PREDICT])
    elif Global_Parameters.GB_MODEL:
        KAGGLE_PREDICTION = pd.DataFrame(DATA_MODEL_GB.MODEL.predict(Global_Data.TEST_DATAFRAME), columns = [Global_Parameters.NAME_DATA_PREDICT])
    elif Global_Parameters.XG_MODEL:
        KAGGLE_PREDICTION = pd.DataFrame(DATA_MODEL_XG.MODEL.predict(Global_Data.TEST_DATAFRAME), columns = [Global_Parameters.NAME_DATA_PREDICT])
    elif Global_Parameters.CB_MODEL:
        KAGGLE_PREDICTION = pd.DataFrame(DATA_MODEL_CB.MODEL.predict(Global_Data.TEST_DATAFRAME), columns = [Global_Parameters.NAME_DATA_PREDICT])
    elif Global_Parameters.LB_MODEL:
        KAGGLE_PREDICTION = pd.DataFrame(DATA_MODEL_LB.MODEL.predict(Global_Data.TEST_DATAFRAME), columns = [Global_Parameters.NAME_DATA_PREDICT])
    Global_Data.data_import(Global_Parameters.NAME_DATA_PREDICT)
    KAGGLE_PREDICTION.index = Global_Data.TEST_DATAFRAME.id
    KAGGLE_PREDICTION.to_csv("kaggle_compet.csv", index_label = "id")

# Model Stacking
else:

    DATA_MODEL_ST = stacking(Data_Model, Global_Parameters, Global_Data,
                              DATA_MODEL_XG.MODEL, DATA_MODEL_CB.MODEL, DATA_MODEL_LB.MODEL)
    
    # Kaggle competition prediction
    KAGGLE_PREDICTION = pd.DataFrame(DATA_MODEL_ST.MODEL.predict(Global_Data.TEST_DATAFRAME), columns = [Global_Parameters.NAME_DATA_PREDICT])
    Global_Data.data_import(Global_Parameters.NAME_DATA_PREDICT)
    KAGGLE_PREDICTION.index = Global_Data.TEST_DATAFRAME.id
    KAGGLE_PREDICTION.to_csv("kaggle_compet.csv", index_label = "id")


#
# Searching for unique combinaison of values

# Brand and model
LEAD_NAME = "brand"
SEARCH_NAME = "model"
UNIQUE_BRAND = pd.unique(Global_Data.TRAIN_DATAFRAME[LEAD_NAME])
UNIQUE_BRAND_MODEL = np.zeros([UNIQUE_BRAND.shape[0],2], dtype = object)
for i, BRAND in enumerate(UNIQUE_BRAND):
    UNIQUE_BRAND_MODEL[i,0] = BRAND
    UNIQUE_BRAND_MODEL[i,1] = pd.unique(Global_Data.TRAIN_DATAFRAME[SEARCH_NAME].loc[Global_Data.TRAIN_DATAFRAME[LEAD_NAME] == BRAND])

with open('./data_linked/unique_brand_model.joblib', 'wb') as f:
    joblib.dump(UNIQUE_BRAND_MODEL, f)

# Model and fuel-type
LEAD_NAME = "model"
SEARCH_NAME = "fuel_type"
UNIQUE_MODEL = pd.unique(Global_Data.TRAIN_DATAFRAME[LEAD_NAME])
UNIQUE_MODEL_FUEL_TYPE = np.zeros([UNIQUE_MODEL.shape[0],2], dtype = object)
for i, MODEL in enumerate(UNIQUE_MODEL):
    UNIQUE_MODEL_FUEL_TYPE[i,0] = MODEL
    UNIQUE_MODEL_FUEL_TYPE[i,1] = pd.unique(Global_Data.TRAIN_DATAFRAME[SEARCH_NAME].loc[Global_Data.TRAIN_DATAFRAME[LEAD_NAME] == MODEL])

with open('./data_linked/unique_model_fuel_type.joblib', 'wb') as f:
    joblib.dump(UNIQUE_MODEL_FUEL_TYPE, f)

# Model and engine
LEAD_NAME = "model"
SEARCH_NAME = "engine"
UNIQUE_MODEL = pd.unique(Global_Data.TRAIN_DATAFRAME[LEAD_NAME])
UNIQUE_MODEL_ENGINE = np.zeros([UNIQUE_MODEL.shape[0],2], dtype = object)
for i, MODEL in enumerate(UNIQUE_MODEL):
    UNIQUE_MODEL_ENGINE[i,0] = MODEL
    UNIQUE_MODEL_ENGINE[i,1] = pd.unique(Global_Data.TRAIN_DATAFRAME[SEARCH_NAME].loc[Global_Data.TRAIN_DATAFRAME[LEAD_NAME] == MODEL])

with open('./data_linked/unique_model_engine.joblib', 'wb') as f:
    joblib.dump(UNIQUE_MODEL_ENGINE, f)

# Model and transmission
LEAD_NAME = "model"
SEARCH_NAME = "transmission"
UNIQUE_MODEL = pd.unique(Global_Data.TRAIN_DATAFRAME[LEAD_NAME])
UNIQUE_MODEL_TRANSMISSION = np.zeros([UNIQUE_MODEL.shape[0],2], dtype = object)
for i, MODEL in enumerate(UNIQUE_MODEL):
    UNIQUE_MODEL_TRANSMISSION[i,0] = MODEL
    UNIQUE_MODEL_TRANSMISSION[i,1] = pd.unique(Global_Data.TRAIN_DATAFRAME[SEARCH_NAME].loc[Global_Data.TRAIN_DATAFRAME[LEAD_NAME] == MODEL])

with open('./data_linked/unique_model_transmission.joblib', 'wb') as f:
    joblib.dump(UNIQUE_MODEL_TRANSMISSION, f)