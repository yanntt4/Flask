# -*- coding: utf-8 -*-
"""
Created on Fri Aug 30 09:38:00 2024

@author: ythiriet
"""

# Global library import
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plot

from sklearn.metrics import matthews_corrcoef
import tensorflow as tf
import xgboost as xgb

import optuna

# Personal library
sys.path.append("./")
from Data_modelling import Data_modelling


# Class definition
class Data_modelling_xgboosting(Data_modelling):
    def __init__(self,num_class):
        super(Data_modelling_xgboosting, self).__init__()

        self.objective_classification='multi:softmax'
        self.objective_regression='reg:linear'
        self.num_class=num_class
        self.learning_rate=0.1237097909626253
        self.max_depth=11
        # self.gamma=0.940859
        self.gamma=0
        self.reg_lambda=5
        self.min_child_weight = 1
        self.early_stopping_rounds=25
        self.eval_metric_classification=['merror','mlogloss']

        self.x_axis = []
        self.results_metric_plot = []


    def xgboosting_modellisation(self, k_folds, REGRESSION, STACKING):

        # Setting the model with parameters
        if REGRESSION:
            
            # Removing early stopping for stacking
            if STACKING:
                self.MODEL = xgb.XGBRegressor(
                    objective=self.objective_regression,
                    learning_rate=self.learning_rate,
                    max_depth=self.max_depth,
                    gamma=self.gamma,
                    reg_lambda=self.reg_lambda,
                    min_child_weight=self.min_child_weight,
                    seed=42)
            
            # No stacking
            else:
                self.MODEL = xgb.XGBRegressor(
                    objective=self.objective_regression,
                    learning_rate=self.learning_rate,
                    max_depth=self.max_depth,
                    gamma=self.gamma,
                    reg_lambda=self.reg_lambda,
                    early_stopping_rounds=self.early_stopping_rounds,
                    min_child_weight=self.min_child_weight,
                    seed=42)
            
        
        else:
            
            # Removing early stopping for stacking
            if STACKING:
                self.MODEL = xgb.XGBClassifier(
                    objective=self.objective_classification,
                    num_class=self.num_class,
                    learning_rate=self.learning_rate,
                    max_depth=self.max_depth,
                    gamma=self.gamma,
                    reg_lambda=self.reg_lambda,
                    eval_metric=self.eval_metric_classification,
                    min_child_weight=self.min_child_weight,
                    seed=42)
            
            # No stacking
            else:
                self.MODEL = xgb.XGBClassifier(
                    objective=self.objective_classification,
                    num_class=self.num_class,
                    learning_rate=self.learning_rate,
                    max_depth=self.max_depth,
                    gamma=self.gamma,
                    reg_lambda=self.reg_lambda,
                    early_stopping_rounds=self.early_stopping_rounds,
                    eval_metric=self.eval_metric_classification,
                    min_child_weight=self.min_child_weight,
                    seed=42)
            
        # Cross validation
        if STACKING == False:
            self.MODEL.fit(self.X_train, self.Y_train,
                            verbose = 1,
                            eval_set = [(self.X_train, self.Y_train),
                                        (self.X_test, self.Y_test)])
        
        # No stacking
        else:
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

            # Preparing evaluation metric plots
            self.results_metric_plot = self.MODEL.evals_result()
            epochs = len(self.results_metric_plot['validation_0']['mlogloss'])
            self.x_axis = range(0, epochs)


    def evaluation_metric_plot_mlogloss(self):

        # xgboost 'mlogloss' plot
        fig, ax = plot.subplots(figsize=(9,5))
        ax.plot(self.x_axis, self.results_metric_plot['validation_0']['mlogloss'], label='Train')
        ax.plot(self.x_axis, self.results_metric_plot['validation_1']['mlogloss'], label='Test')
        ax.legend()
        plot.ylabel('mlogloss')
        plot.title('GridSearchCV XGBoost mlogloss')
        plot.show()


    def evaluation_metric_plot_merror(self):

        # xgboost 'merror' plot
        fig, ax = plot.subplots(figsize=(9,5))
        ax.plot(self.x_axis, self.results_metric_plot['validation_0']['merror'], label='Train')
        ax.plot(self.x_axis, self.results_metric_plot['validation_1']['merror'], label='Test')
        ax.legend()
        plot.ylabel('merror')
        plot.title('GridSearchCV XGBoost merror')
        plot.show()


    def feature_importance_plot(self):

        fig, ax = plot.subplots(figsize=(9,5))
        xgb.plot_importance(self.MODEL, ax=ax)
        plot.show()


# Function to create/optimize xgboosting model
def xgboosting(Data_Model, Global_Parameters, Global_Data):
    DATA_MODEL_XG = Data_modelling_xgboosting(pd.unique(Global_Data.TRAIN_DATAFRAME[Global_Parameters.NAME_DATA_PREDICT]).shape[0])
    
    # Using split create previously
    DATA_MODEL_XG.X_train = Data_Model.X_train
    DATA_MODEL_XG.Y_train = Data_Model.Y_train
    DATA_MODEL_XG.X_test = Data_Model.X_test
    DATA_MODEL_XG.Y_test = Data_Model.Y_test
    DATA_MODEL_XG.MODEL_NAME = "XG Boosting"
    
    #
    # Building a Gradient boosting Model with adjusted parameters
    
    # Regression
    if Global_Parameters.REGRESSION:
        
        def build_model_XG(objective='reg:linear', learning_rate=0.1, max_depth=5,
                           gamma=0, reg_lambda=1, early_stopping_rounds=25,
                           min_child_weight=1):
        
            MODEL_XG = xgb.XGBRegressor(
                objective=objective,
                learning_rate=learning_rate,
                max_depth=max_depth,
                gamma=gamma,
                reg_lambda=reg_lambda,
                early_stopping_rounds=early_stopping_rounds,
                min_child_weight = min_child_weight,
                seed=42)
            
            return MODEL_XG
    
    # Classification
    else:
    
        def build_model_XG(objective='multi:softmax', num_class=16, learning_rate=0.1, max_depth=5,
                           gamma=0, reg_lambda=1, early_stopping_rounds=25,
                           eval_metric = ['merror','mlogloss'], min_child_weight=1):
    
            MODEL_XG = xgb.XGBClassifier(
                objective=objective,
                num_class=num_class,
                learning_rate=learning_rate,
                max_depth=max_depth,
                gamma=gamma,
                reg_lambda=reg_lambda,
                early_stopping_rounds=early_stopping_rounds,
                eval_metric=eval_metric,
                min_child_weight = min_child_weight,
                seed=42)
            
            return MODEL_XG
    
        
    # Searching for Optimized Hyperparameters
    if Global_Parameters.XG_MODEL_OPTI:
        
        # Regression
        if Global_Parameters.REGRESSION:
    
            # Building function to minimize/maximise
            def objective_XG(trial):
                params = {'learning_rate': trial.suggest_float('learning_rate', 0.01, 0.5),
                          'max_depth': trial.suggest_int('max_depth', 1, 15),
                          'gamma': trial.suggest_float('gamma', 0, 1),
                          'min_child_weight': trial.suggest_int('min_child_weight', 1, 50),
                          'reg_lambda': trial.suggest_int('reg_lambda', 1, 10)}
        
                MODEL_XG = build_model_XG(**params)
                MODEL_XG.fit(DATA_MODEL_XG.X_train, DATA_MODEL_XG.Y_train,
                               verbose = 1,
                               eval_set = [(DATA_MODEL_XG.X_train, DATA_MODEL_XG.Y_train),
                                           (DATA_MODEL_XG.X_test, DATA_MODEL_XG.Y_test)])
                
                MSLE = tf.keras.losses.MSLE(DATA_MODEL_XG.Y_test, MODEL_XG.predict(DATA_MODEL_XG.X_test))
                
                # Exit
                return MSLE
            
        # Classification
        else:
            
            # Building function to minimize/maximise
            def objective_XG(trial):
                params = {'learning_rate': trial.suggest_float('learning_rate', 0.01, 0.5),
                          'max_depth': trial.suggest_int('max_depth', 1, 15),
                          'gamma': trial.suggest_float('gamma', 0, 1),
                          'min_child_weight': trial.suggest_int('min_child_weight', 1, 50),
                          'reg_lambda': trial.suggest_int('reg_lambda', 1, 10)}
        
                MODEL_XG = build_model_XG(**params)
                MODEL_XG.fit(DATA_MODEL_XG.X_train, DATA_MODEL_XG.Y_train,
                               verbose = 1,
                               eval_set = [(DATA_MODEL_XG.X_train, DATA_MODEL_XG.Y_train),
                                           (DATA_MODEL_XG.X_test, DATA_MODEL_XG.Y_test)])
                
                MATTHEWS_CORRCOEF = matthews_corrcoef(DATA_MODEL_XG.Y_test, MODEL_XG.predict(DATA_MODEL_XG.X_test))
                
                # Exit
                return MATTHEWS_CORRCOEF
        
    
        # Search for best hyperparameters
        
        # Regression
        if Global_Parameters.REGRESSION:
            study = optuna.create_study(direction='minimize')
        else:
            study = optuna.create_study(direction='maximize')
        study.optimize(objective_XG, n_trials=Global_Parameters.XG_MODEL_TRIAL, catch=(ValueError,))
        
        # Saving and using best hyperparameters
        BEST_PARAMS_XG = np.zeros([1], dtype=object)
        BEST_PARAMS_XG[0] = study.best_params
        DATA_MODEL_XG.learning_rate = float(BEST_PARAMS_XG[0].get("learning_rate"))
        DATA_MODEL_XG.max_depth = int(BEST_PARAMS_XG[0].get("max_depth"))
        DATA_MODEL_XG.gamma = int(BEST_PARAMS_XG[0].get("gamma"))
        DATA_MODEL_XG.min_child_weights = int(BEST_PARAMS_XG[0].get("min_child_weight"))
        DATA_MODEL_XG.reg_lambda = int(BEST_PARAMS_XG[0].get("reg_lambda"))
    
    # Creating and fitting xgboosting model
    DATA_MODEL_XG.xgboosting_modellisation(
        Global_Parameters.k_folds, Global_Parameters.REGRESSION, Global_Parameters.STACKING)
    
    #
    # Result analysis
    
    # No stacking
    if Global_Parameters.STACKING == False:
        DATA_MODEL_XG.feature_importance_plot()
        
        # Classification
        if Global_Parameters.REGRESSION == False:
            DATA_MODEL_XG.evaluation_metric_plot_mlogloss()
            DATA_MODEL_XG.evaluation_metric_plot_merror()
            
            DATA_MODEL_XG.result_plot_classification()
            DATA_MODEL_XG.result_report_classification_calculation()
            DATA_MODEL_XG.result_report_classification_print()
            DATA_MODEL_XG.result_report_classification_plot()
            
        # Regression
        else:  
            DATA_MODEL_XG.result_plot_regression(Global_Data.TRAIN_DATAFRAME, Global_Parameters.NAME_DATA_PREDICT)
            DATA_MODEL_XG.result_report_regression_calculation()
            DATA_MODEL_XG.result_report_regression_print()
            DATA_MODEL_XG.result_report_regression_plot()
            
            DATA_MODEL_XG.extract_max_diff_regression()
    
    # Exit
    return DATA_MODEL_XG