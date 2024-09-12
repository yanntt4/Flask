# -*- coding: utf-8 -*-
"""
Created on Fri Aug 30 16:42:32 2024

@author: ythiriet
"""

# Global library import
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plot

from sklearn.ensemble import RandomForestRegressor, RandomForestClassifier
from sklearn.ensemble import StackingRegressor, StackingClassifier

# Personal library
sys.path.append("./")
from Data_modelling import Data_modelling


# Class definition
class Data_modelling_stacking(Data_modelling):
    def __init__(self, num_class, XG_MODEL, CB_MODEL, LB_MODEL):
        super(Data_modelling_stacking, self).__init__()
        
        self.estimators = [
            ('XG', XG_MODEL),
            ('CB', CB_MODEL),
            ('LB', LB_MODEL)]


    def stacking_modellisation(self, k_folds, REGRESSION, STACKING):

        # Setting the model with parameters
        if REGRESSION:
            self.MODEL = StackingRegressor(
                estimators=self.estimators, final_estimator=RandomForestRegressor())
        
        else:
            self.MODEL = StackingClassifier(
                estimators=self.estimators, final_estimator=RandomForestClassifier())
        
        # Fitting results
        self.MODEL.fit(self.X_train, self.Y_train)
 
        # Predicting results
        self.Y_predict = self.MODEL.predict(self.X_test)
        Y_test = np.squeeze(self.Y_test.to_numpy())

        # Percentage calculation for correct prediction
        if REGRESSION:
            self.AVERAGE_DIFFERENCE = np.mean(abs(self.Y_predict - Y_test))
            print(f"\n Moyenne des différences : {round(self.AVERAGE_DIFFERENCE,2)} €")
            self.PERCENTAGE_AVERAGE_DIFFERENCE = 100*self.AVERAGE_DIFFERENCE / np.mean(Y_test)
            print(f"\n Pourcentage de différence : {round(self.PERCENTAGE_AVERAGE_DIFFERENCE,2)} %")
        else:
            self.Y_PREDICT_PROBA = self.MODEL.predict_proba(self.X_test)
            self.NB_CORRECT_PREDICTION = np.count_nonzero(
                Y_test.astype(int) - self.Y_predict.astype(int))
            self.PERCENTAGE_CORRECT_PREDICTION = (1 -
                self.NB_CORRECT_PREDICTION / Y_test.shape[0])
            print(f"\n Pourcentage de predictions correctes {self.MODEL_NAME} : {100*round(self.PERCENTAGE_CORRECT_PREDICTION,5)} %")


# Function to create/optimize xgboosting model
def stacking(Data_Model, Global_Parameters, Global_Data,
             XG_MODEL, CB_MODEL, LB_MODEL):
    DATA_MODEL_ST = Data_modelling_stacking(pd.unique(Global_Data.TRAIN_DATAFRAME[Global_Parameters.NAME_DATA_PREDICT]).shape[0],
                                            XG_MODEL, CB_MODEL, LB_MODEL)
    
    # Using split create previously
    DATA_MODEL_ST.X_train = Data_Model.X_train
    DATA_MODEL_ST.Y_train = Data_Model.Y_train
    DATA_MODEL_ST.X_test = Data_Model.X_test
    DATA_MODEL_ST.Y_test = Data_Model.Y_test
    DATA_MODEL_ST.MODEL_NAME = "Stacking"

    # Creating and fitting xgboosting model
    DATA_MODEL_ST.stacking_modellisation(
        Global_Parameters.k_folds, Global_Parameters.REGRESSION, Global_Parameters.STACKING)
    
    #
    # Result analysis
    
    # Classification
    if Global_Parameters.REGRESSION == False:
    
        DATA_MODEL_ST.result_plot_classification()
        DATA_MODEL_ST.result_report_classification_calculation()
        DATA_MODEL_ST.result_report_classification_print()
        DATA_MODEL_ST.result_report_classification_plot()
        
    # Regression
    else:  
        DATA_MODEL_ST.result_plot_regression(Global_Data.TRAIN_DATAFRAME, Global_Parameters.NAME_DATA_PREDICT)
        DATA_MODEL_ST.result_report_regression_calculation()
        DATA_MODEL_ST.result_report_regression_print()
        DATA_MODEL_ST.result_report_regression_plot()
        
        DATA_MODEL_ST.extract_max_diff_regression()


    # Exit
    return DATA_MODEL_ST