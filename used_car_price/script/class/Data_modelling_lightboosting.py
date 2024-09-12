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
import lightgbm

import optuna

# Personal library
sys.path.append("./")
from Data_modelling import Data_modelling


# Class definition
class Data_modelling_lightboosting(Data_modelling):
    def __init__(self,num_class):
        super(Data_modelling_lightboosting, self).__init__()
        
        self.learning_rate=0.5
        self.objective_regression='regression'
        self.objective_classification='binary' #multiclass
        self.min_split_gain = 0.5273582861187036
        self.num_leaves=785
        self.max_depth=7
        self.min_child_samples=98
        self.pos_bagging_fraction=0.20229832313174725
        self.neg_bagging_fraction=0.34192297208528094
        self.reg_alpha=0.023948778772931792
        self.reg_lambda=0.010117261678790731
        self.random_state = 42

        self.x_axis = []
        self.results_metric_plot = []


    def lightboosting_modellisation(self, k_folds, REGRESSION, STACKING):

        # Setting the model with parameters
        if REGRESSION:
            self.MODEL = lightgbm.LGBMRegressor(
                objective=self.objective_regression,
                learning_rate=self.learning_rate,
                min_split_gain=self.min_split_gain,
                num_leaves=self.num_leaves,
                max_depth=self.max_depth,
                min_child_samples=self.min_child_samples,
                pos_bagging_fraction=self.pos_bagging_fraction,
                neg_bagging_fraction=self.neg_bagging_fraction,
                reg_alpha=self.reg_alpha,
                reg_lambda=self.reg_lambda,
                random_state=self.random_state)
        
        else:
            self.MODEL = lightgbm.LGBMClassifier(
                objective=self.objective_classification,
                learning_rate=self.learning_rate,
                min_split_gain=self.min_split_gain,
                num_leaves=self.num_leaves,
                max_depth=self.max_depth,
                min_child_samples=self.min_child_samples,
                pos_bagging_fraction=self.pos_bagging_fraction,
                neg_bagging_fraction=self.neg_bagging_fraction,
                reg_alpha=self.reg_alpha,
                reg_lambda=self.reg_lambda,
                random_state=self.random_state)
            
        # Cross validation with early stopping
        if STACKING:
            self.MODEL.fit(self.X_train, self.Y_train)
        else:
            callbacks = [lightgbm.early_stopping(50, verbose=0), lightgbm.log_evaluation(period=0)]
            self.MODEL.fit(self.X_train, self.Y_train,
                            eval_set=[(self.X_train, self.Y_train), (self.X_test, self.Y_test)],
                            eval_names=['train', 'test'],
                            eval_metric='average_precision',
                            callbacks=callbacks)
        
 
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
def lightboosting(Data_Model, Global_Parameters, Global_Data):
    DATA_MODEL_LB = Data_modelling_lightboosting(pd.unique(Global_Data.TRAIN_DATAFRAME[Global_Parameters.NAME_DATA_PREDICT]).shape[0])
    
    # Using split create previously
    DATA_MODEL_LB.X_train = Data_Model.X_train
    DATA_MODEL_LB.Y_train = Data_Model.Y_train
    DATA_MODEL_LB.X_test = Data_Model.X_test
    DATA_MODEL_LB.Y_test = Data_Model.Y_test
    DATA_MODEL_LB.MODEL_NAME = "Lightboosting"
    
    #
    # Building a Gradient boosting Model with adjusted parameters
    
    # Regression
    if Global_Parameters.REGRESSION:
        
        def build_model_LB(learning_rate=0.1, min_split_gain = 1, num_leaves=100,
                           max_depth=5, min_child_samples=20, pos_bagging_fraction=0.5,
                           neg_bagging_fraction=0.5, reg_alpha=0.01, reg_lambda=0.01):
        
            MODEL_LB = lightgbm.LGBMRegressor(
                objective='regression',
                learning_rate=learning_rate,
                min_split_gain=min_split_gain,
                num_leaves=num_leaves,
                max_depth=max_depth,
                min_child_samples=min_child_samples,
                pos_bagging_fraction=pos_bagging_fraction,
                neg_bagging_fraction=neg_bagging_fraction,
                reg_alpha=reg_alpha,
                reg_lambda=reg_lambda,
                random_state=42)
            
            return MODEL_LB
    
    # Classification
    else:
    
        def build_model_LB(learning_rate=0.1, min_split_gain = 1, num_leaves=100,
                           max_depth=5, min_child_samples=20, pos_bagging_fraction=0.5,
                           neg_bagging_fraction=0.5, reg_alpha=0.01, reg_lambda=0.01):
    
            MODEL_LB = lightgbm.LGBMClassifier(
                objective='binary',
                learning_rate=learning_rate,
                min_split_gain=min_split_gain,
                num_leaves=num_leaves,
                max_depth=max_depth,
                min_child_samples=min_child_samples,
                pos_bagging_fraction=pos_bagging_fraction,
                neg_bagging_fraction=neg_bagging_fraction,
                reg_alpha=reg_alpha,
                reg_lambda=reg_lambda,
                random_state=42)
            
            return MODEL_LB
    
        
    # Searching for Optimized Hyperparameters
    if Global_Parameters.LB_MODEL_OPTI:
        
        # Regression
        if Global_Parameters.REGRESSION:
    
            # Building function to minimize/maximise
            def objective_LB(trial):
                params = {'learning_rate': trial.suggest_float('learning_rate',0.01, 0.5),
                          'min_split_gain': trial.suggest_float('min_split_gain',0.00001, 2),
                          'num_leaves': trial.suggest_int('num_leaves',2, 1024),
                          'max_depth': trial.suggest_int('max_depth',1, 15),
                          'min_child_samples': trial.suggest_int('min_child_samples',2, 100),
                          'pos_bagging_fraction': trial.suggest_float('pos_bagging_fraction',0, 1),
                          'neg_bagging_fraction': trial.suggest_float('neg_bagging_fraction',0, 1),
                          'reg_alpha': trial.suggest_float('reg_alpha',0.00001, 0.1),
                          'reg_lambda': trial.suggest_float('reg_lambda',0.00001, 0.1)}
        
                MODEL_LB = build_model_LB(**params)
                callbacks = [lightgbm.early_stopping(50, verbose=0), lightgbm.log_evaluation(period=0)]
                MODEL_LB.fit(DATA_MODEL_LB.X_train, DATA_MODEL_LB.Y_train,
                             eval_set=[(DATA_MODEL_LB.X_train, DATA_MODEL_LB.Y_train), (DATA_MODEL_LB.X_test, DATA_MODEL_LB.Y_test)],
                             eval_names=['train', 'test'],
                             eval_metric='average_precision',
                             callbacks=callbacks)
                
                MSLE = tf.keras.losses.MSLE(DATA_MODEL_LB.Y_test, MODEL_LB.predict(DATA_MODEL_LB.X_test))
                
                # Exit
                return MSLE
            
        # Classification
        else:
            
            # Building function to minimize/maximise
            def objective_LB(trial):
                params = {'learning_rate': trial.suggest_float('learning_rate',0.01, 0.5),
                          'min_split_gain': trial.suggest_float('min_split_gain',0.00001, 2),
                          'num_leaves': trial.suggest_int('num_leaves',2, 1024),
                          'max_depth': trial.suggest_int('max_depth',1, 15),
                          'min_child_samples': trial.suggest_int('min_child_samples',2, 100),
                          'pos_bagging_fraction': trial.suggest_float('pos_bagging_fraction',0, 1),
                          'neg_bagging_fraction': trial.suggest_float('neg_bagging_fraction',0, 1),
                          'reg_alpha': trial.suggest_float('reg_alpha',0.00001, 0.1),
                          'reg_lambda': trial.suggest_float('reg_lambda',0.00001, 0.1)}
        
                MODEL_LB = build_model_LB(**params)
                callbacks = [lightgbm.early_stopping(50, verbose=0), lightgbm.log_evaluation(period=0)]
                MODEL_LB.fit(DATA_MODEL_LB.X_train, DATA_MODEL_LB.Y_train,
                             eval_set=[(DATA_MODEL_LB.X_train, DATA_MODEL_LB.Y_train), (DATA_MODEL_LB.X_test, DATA_MODEL_LB.Y_test)],
                             eval_names=['train', 'test'],
                             eval_metric='average_precision',
                             callbacks=callbacks)
                
                MATTHEWS_CORRCOEF = matthews_corrcoef(DATA_MODEL_LB.Y_test, MODEL_LB.predict(DATA_MODEL_LB.X_test))
                
                # Exit
                return MATTHEWS_CORRCOEF
        
    
        # Search for best hyperparameters
        
        # Regression
        if Global_Parameters.REGRESSION:
            study = optuna.create_study(direction='minimize')
        else:
            study = optuna.create_study(direction='maximize')
        study.optimize(objective_LB, n_trials=Global_Parameters.LB_MODEL_TRIAL, catch=(ValueError,))
        
        # Saving and using best hyperparameters
        BEST_PARAMS_LB = np.zeros([1], dtype=object)
        BEST_PARAMS_LB[0] = study.best_params
        DATA_MODEL_LB.learning_rate = float(BEST_PARAMS_LB[0].get("learning_rate"))
        DATA_MODEL_LB.min_split_gain = float(BEST_PARAMS_LB[0].get("min_split_gain")),
        DATA_MODEL_LB.num_leaves = int(BEST_PARAMS_LB[0].get("num_leaves")),
        DATA_MODEL_LB.max_depth = int(BEST_PARAMS_LB[0].get("max_depth")),
        DATA_MODEL_LB.min_child_samples = int(BEST_PARAMS_LB[0].get("min_child_samples")),
        DATA_MODEL_LB.pos_bagging_fraction = float(BEST_PARAMS_LB[0].get("pos_bagging_fraction")),
        DATA_MODEL_LB.neg_bagging_fraction = float(BEST_PARAMS_LB[0].get("neg_bagging_fraction")),
        DATA_MODEL_LB.reg_alpha = float(BEST_PARAMS_LB[0].get("reg_alpha")),
        DATA_MODEL_LB.reg_lambda = float(BEST_PARAMS_LB[0].get("reg_lambda"))
    
    # Creating and fitting xgboosting model
    DATA_MODEL_LB.lightboosting_modellisation(
        Global_Parameters.k_folds, Global_Parameters.REGRESSION, Global_Parameters.STACKING)
    
    #
    # Result analysis
    
    if Global_Parameters.STACKING == False:
    
        # Classification
        if Global_Parameters.REGRESSION == False:
        
            DATA_MODEL_LB.result_plot_classification()
            DATA_MODEL_LB.result_report_classification_calculation()
            DATA_MODEL_LB.result_report_classification_print()
            DATA_MODEL_LB.result_report_classification_plot()
            
        # Regression
        else:  
            DATA_MODEL_LB.result_plot_regression(Global_Data.TRAIN_DATAFRAME, Global_Parameters.NAME_DATA_PREDICT)
            DATA_MODEL_LB.result_report_regression_calculation()
            DATA_MODEL_LB.result_report_regression_print()
            DATA_MODEL_LB.result_report_regression_plot()
            
            DATA_MODEL_LB.extract_max_diff_regression()


    # Exit
    return DATA_MODEL_LB