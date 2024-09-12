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

from sklearn.metrics import matthews_corrcoef
import tensorflow as tf
import catboost

import optuna

# Personal library
sys.path.append("./")
from Data_modelling import Data_modelling


# Class definition
class Data_modelling_catboosting(Data_modelling):
    def __init__(self,num_class):
        super(Data_modelling_catboosting, self).__init__()
        
        self.learning_rate=0.2763960654116985
        self.loss_function_regression='RMSE'
        self.loss_function_classification='CrossEntropy'
        self.depth = 3
        self.n_estimators=896
        self.early_stopping_rounds=25
        self.l2_leaf_reg = 1

        self.x_axis = []
        self.results_metric_plot = []


    def catboosting_modellisation(self, k_folds, REGRESSION, STACKING):

        # Setting the model with parameters
        if REGRESSION:
            self.MODEL = catboost.CatBoostRegressor(
                learning_rate=self.learning_rate,
                loss_function=self.loss_function_regression,
                depth=self.depth,
                n_estimators=self.n_estimators,
                l2_leaf_reg=self.l2_leaf_reg,
                random_seed=42)
        
        else:
            self.MODEL = catboost.CatBoostClassifier(
                learning_rate=self.learning_rate,
                loss_function=self.loss_function_classification,
                depth=self.depth,
                n_estimators=self.n_estimators,
                l2_leaf_reg=self.l2_leaf_reg,
                random_seed=42)
            
        # Cross validation with early stopping
        if STACKING:
            self.MODEL.fit(self.X_train, self.Y_train)
        else:
            self.MODEL.fit(self.X_train, self.Y_train, verbose = 1,
                           eval_set = catboost.Pool(self.X_test, self.Y_test),
                           early_stopping_rounds=self.early_stopping_rounds)
        
 
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
def catboosting(Data_Model, Global_Parameters, Global_Data):
    DATA_MODEL_CB = Data_modelling_catboosting(pd.unique(Global_Data.TRAIN_DATAFRAME[Global_Parameters.NAME_DATA_PREDICT]).shape[0])
    
    # Using split create previously
    DATA_MODEL_CB.X_train = Data_Model.X_train
    DATA_MODEL_CB.Y_train = Data_Model.Y_train
    DATA_MODEL_CB.X_test = Data_Model.X_test
    DATA_MODEL_CB.Y_test = Data_Model.Y_test
    DATA_MODEL_CB.MODEL_NAME = "Catboosting"
    
    #
    # Building a Gradient boosting Model with adjusted parameters
    
    # Regression
    if Global_Parameters.REGRESSION:
        
        def build_model_CB(learning_rate=0.1, loss_function='RMSE', depth = 2,
                           n_estimators=5, l2_leaf_reg = 1):
        
            MODEL_CB = catboost.CatBoostRegressor(
                learning_rate=learning_rate,
                loss_function=loss_function,
                depth=depth,
                n_estimators=n_estimators,
                l2_leaf_reg=l2_leaf_reg,
                random_seed=42)
            
            return MODEL_CB
    
    # Classification
    else:
    
        def build_model_CB(learning_rate=0.1, loss_function='CrossEntropy', depth = 2,
                           n_estimators=5, l2_leaf_reg = 1):
    
            MODEL_CB = catboost.CatBoostClassifier(
                learning_rate=learning_rate,
                loss_function=loss_function,
                depth=depth,
                n_estimators=n_estimators,
                l2_leaf_reg=l2_leaf_reg,
                random_seed=42)
            
            return MODEL_CB
    
        
    # Searching for Optimized Hyperparameters
    if Global_Parameters.CB_MODEL_OPTI:
        
        # Regression
        if Global_Parameters.REGRESSION:
    
            # Building function to minimize/maximise
            def objective_CB(trial):
                params = {'learning_rate': trial.suggest_float('learning_rate', 0.01, 0.9),
                          'depth': trial.suggest_int('depth', 1, 16),
                          'n_estimators': trial.suggest_int('n_estimators', 1, 1000),
                          'l2_leaf_reg': trial.suggest_float('l2_leaf_reg', 0.01, 5)}
        
                MODEL_CB = build_model_CB(**params)
                MODEL_CB.fit(DATA_MODEL_CB.X_train, DATA_MODEL_CB.Y_train,verbose = 1,
                             eval_set = catboost.Pool(DATA_MODEL_CB.X_test, DATA_MODEL_CB.Y_test),
                             early_stopping_rounds=DATA_MODEL_CB.early_stopping_rounds)
                
                MSLE = tf.keras.losses.MSLE(DATA_MODEL_CB.Y_test, MODEL_CB.predict(DATA_MODEL_CB.X_test))
                
                # Exit
                return MSLE
            
        # Classification
        else:
            
            # Building function to minimize/maximise
            def objective_CB(trial):
                params = {'learning_rate': trial.suggest_float('learning_rate', 0.01, 0.5),
                          'depth': trial.suggest_int('depth', 1, 15),
                          'n_estimators': trial.suggest_int('n_estimators', 1, 1000),
                          'l2_leaf_reg': trial.suggest_float('l2_leaf_reg', 0.2, 5)}
        
                MODEL_CB = build_model_CB(**params)
                MODEL_CB.fit(DATA_MODEL_CB.X_train, DATA_MODEL_CB.Y_train,verbose = 1,
                             eval_set = catboost.Pool(DATA_MODEL_CB.X_test, DATA_MODEL_CB.Y_test),
                             early_stopping_rounds=DATA_MODEL_CB.early_stopping_rounds)
                
                MATTHEWS_CORRCOEF = matthews_corrcoef(DATA_MODEL_CB.Y_test, MODEL_CB.predict(DATA_MODEL_CB.X_test))
                
                # Exit
                return MATTHEWS_CORRCOEF
        
    
        # Search for best hyperparameters
        
        # Regression
        if Global_Parameters.REGRESSION:
            study = optuna.create_study(direction='minimize')
        else:
            study = optuna.create_study(direction='maximize')
        study.optimize(objective_CB, n_trials=Global_Parameters.CB_MODEL_TRIAL, catch=(ValueError,))
        
        # Saving and using best hyperparameters
        BEST_PARAMS_CB = np.zeros([1], dtype=object)
        BEST_PARAMS_CB[0] = study.best_params
        DATA_MODEL_CB.learning_rate = float(BEST_PARAMS_CB[0].get("learning_rate"))
        DATA_MODEL_CB.depth = int(BEST_PARAMS_CB[0].get("depth"))
        DATA_MODEL_CB.n_estimators = int(BEST_PARAMS_CB[0].get("n_estimators"))
        DATA_MODEL_CB.l2_leaf_reg = int(BEST_PARAMS_CB[0].get("l2_leaf_reg"))
    
    # Creating and fitting xgboosting model
    DATA_MODEL_CB.catboosting_modellisation(
        Global_Parameters.k_folds, Global_Parameters.REGRESSION, Global_Parameters.STACKING)
    
    #
    # Result analysis
    
    # No stacking
    if Global_Parameters.STACKING == False:
    
        # Classification
        if Global_Parameters.REGRESSION == False:
        
            DATA_MODEL_CB.result_plot_classification()
            DATA_MODEL_CB.result_report_classification_calculation()
            DATA_MODEL_CB.result_report_classification_print()
            DATA_MODEL_CB.result_report_classification_plot()
            
        # Regression
        else:  
            DATA_MODEL_CB.result_plot_regression(Global_Data.TRAIN_DATAFRAME, Global_Parameters.NAME_DATA_PREDICT)
            DATA_MODEL_CB.result_report_regression_calculation()
            DATA_MODEL_CB.result_report_regression_print()
            DATA_MODEL_CB.result_report_regression_plot()
            
            DATA_MODEL_CB.extract_max_diff_regression()


    # Exit
    return DATA_MODEL_CB