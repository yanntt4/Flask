# -*- coding: utf-8 -*-
"""
Created on Thu Aug 29 16:10:14 2024

@author: ythiriet
"""

# Global library import
import sys
import numpy as np
import matplotlib.pyplot as plot

from sklearn.ensemble import GradientBoostingClassifier, GradientBoostingRegressor
from sklearn.model_selection import cross_val_score, KFold
from sklearn.metrics import mean_squared_error

from sklearn.metrics import matthews_corrcoef

import optuna

# Personal library
sys.path.append("./")
from Data_modelling import Data_modelling


# Class definition
class Data_modelling_gradient_boosting(Data_modelling):
    def __init__(self):
        super(Data_modelling_gradient_boosting, self).__init__()

        self.learning_rate = 0.4968286095170373
        self.Nb_Tree = 192
        self.min_samples_leaf = 37
        self.min_samples_split = 35
        self.min_weight_fraction_leaf = 0.00018800850129667424
        self.max_depth = 24
        self.validation_fraction = 0.1   # Early Stopping
        self.n_iter_no_change = 40   # Early Stopping
        # self.n_iter_no_change = 10   # Early Stopping
        self.train_errors = []   # Early Stopping
        self.test_errors = []   # Early Stopping


    # Model creation and fitting
    def gradient_boosting_modellisation(self, N_SPLIT, REGRESSION):

        # Setting the model with parameters
        if REGRESSION:
            self.MODEL = GradientBoostingRegressor(
                learning_rate=self.learning_rate,
                n_estimators=self.Nb_Tree,
                min_samples_leaf=self.min_samples_leaf,
                min_samples_split=self.min_samples_split,
                max_depth=self.max_depth,
                min_weight_fraction_leaf=self.min_weight_fraction_leaf,
                verbose=2,
                random_state=0,
                validation_fraction=self.validation_fraction,
                n_iter_no_change=self.n_iter_no_change)

        else:
            self.MODEL = GradientBoostingClassifier(
                learning_rate=self.learning_rate,
                n_estimators=self.Nb_Tree,
                min_samples_leaf=self.min_samples_leaf,
                min_samples_split=self.min_samples_split,
                max_depth=self.max_depth,
                min_weight_fraction_leaf=self.min_weight_fraction_leaf,
                verbose=2,
                random_state=0,
                validation_fraction=self.validation_fraction,
                n_iter_no_change=self.n_iter_no_change)

        # Init
        k_folds = KFold(n_splits=N_SPLIT)

        # Cross validation
        self.score = cross_val_score(self.MODEL, self.X_train, self.Y_train, cv=k_folds)
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


    # Plotting evolution of validation errors during training
    def plot_training_validation_error(self):

        for i, (train_pred, test_pred) in enumerate(
            zip(
                self.MODEL.staged_predict(self.X_train),
                self.MODEL.staged_predict(self.X_test),
            )
        ):

            if isinstance(self.Y_train.iloc[0], bool):
                self.train_errors.append(mean_squared_error(
                    self.Y_train, train_pred))
                self.test_errors.append(mean_squared_error(
                    self.Y_test, test_pred))
            else:
                self.train_errors.append(mean_squared_error(
                    self.Y_train.astype(int), train_pred.astype(int)))
                self.test_errors.append(mean_squared_error(
                    self.Y_test.astype(int), test_pred.astype(int)))


        fig, ax = plot.subplots(ncols=2, figsize=(12, 4))

        ax[0].plot(self.train_errors, label="Gradient Boosting with Early Stopping")
        ax[0].set_xlabel("Boosting Iterations")
        ax[0].set_ylabel("MSE (Training)")
        ax[0].set_yscale("log")
        ax[0].legend()
        ax[0].set_title("Training Error")

        ax[1].plot(self.test_errors, label="Gradient Boosting with Early Stopping")
        ax[1].set_xlabel("Boosting Iterations")
        ax[1].set_ylabel("MSE (Validation)")
        ax[1].set_yscale("log")
        ax[1].legend()
        ax[1].set_title("Validation Error")


# Function to create/optimize Gradient Boosting model
def gradient_boosting(Data_Model, Global_Parameters, Global_Data):
    DATA_MODEL_GB = Data_modelling_gradient_boosting()
    DATA_MODEL_GB.X_train = Data_Model.X_train
    DATA_MODEL_GB.Y_train = Data_Model.Y_train
    DATA_MODEL_GB.X_test = Data_Model.X_test
    DATA_MODEL_GB.Y_test = Data_Model.Y_test
    DATA_MODEL_GB.MODEL_NAME = "Gradient Boosting"
    
    #
    # Building a Gradient Boosting Model with adjusted parameters
    
    # Regression
    if Global_Parameters.REGRESSION:
        
        def build_model_GB(
                learning_rate=0.1,
                Nb_Tree=1,
                min_samples_split=10,
                min_samples_leaf=2,
                min_weight_fraction_leaf=0.5,
                max_depth=2):
    
            MODEL_GB = GradientBoostingRegressor(
                learning_rate=learning_rate,
                n_estimators=Nb_Tree,
                min_samples_leaf=min_samples_leaf,
                min_samples_split=min_samples_split,
                random_state=0,
                max_depth=max_depth,
                min_weight_fraction_leaf=min_weight_fraction_leaf)
    
            return MODEL_GB
    
    # Classification
    else:
        
        def build_model_GB(
                learning_rate=0.1,
                Nb_Tree=1,
                min_samples_split=10,
                min_samples_leaf=2,
                min_weight_fraction_leaf=0.5,
                max_depth=2):
    
            MODEL_GB = GradientBoostingClassifier(
                learning_rate=learning_rate,
                n_estimators=Nb_Tree,
                min_samples_leaf=min_samples_leaf,
                min_samples_split=min_samples_split,
                random_state=0,
                max_depth=max_depth,
                min_weight_fraction_leaf=min_weight_fraction_leaf)
    
            return MODEL_GB
    
    # Searching for Optimized Hyperparameters
    if Global_Parameters.GB_MODEL_OPTI:
        
        # Regression
        if Global_Parameters.REGRESSION:
    
            # Building function to minimize/maximise
            def objective_GB(trial):
                params = {'learning_rate': trial.suggest_float('learning_rate', 0.005, 0.5),
                          'Nb_Tree': trial.suggest_int('Nb_Tree', 2, 200),
                          'min_samples_split': trial.suggest_int('min_samples_split', 2, 50),
                          'min_samples_leaf': trial.suggest_int('min_samples_leaf', 2, 50),
                          'min_weight_fraction_leaf': trial.suggest_float('min_weight_fraction_leaf', 0, 0.5),
                          'max_depth': trial.suggest_int('max_depth', 2, 50)}
        
                MODEL_GB = build_model_GB(**params)
                MODEL_GB.fit(Data_Model.X_train, Data_Model.Y_train)
                scores = cross_val_score(MODEL_GB, Data_Model.X_test, Data_Model.Y_test, cv=Global_Parameters.k_folds)
                prediction_score = MODEL_GB.score(Data_Model.X_test, Data_Model.Y_test)
        
                return (4*np.mean(scores) + prediction_score)/5
            
            
        # Classification
        else:
            
            # Building function to minimize/maximise
            def objective_GB(trial):
                params = {'learning_rate': trial.suggest_float('learning_rate', 0.005, 0.5),
                          'Nb_Tree': trial.suggest_int('Nb_Tree', 2, 200),
                          'min_samples_split': trial.suggest_int('min_samples_split', 2, 50),
                          'min_samples_leaf': trial.suggest_int('min_samples_leaf', 2, 50),
                          'min_weight_fraction_leaf': trial.suggest_float('min_weight_fraction_leaf', 0, 0.5),
                          'max_depth': trial.suggest_int('max_depth', 2, 50)}
        
                MODEL_GB = build_model_GB(**params)
                MODEL_GB.fit(Data_Model.X_train, Data_Model.Y_train)
            
                MATTHEWS_CORRCOEF = matthews_corrcoef(DATA_MODEL_GB.Y_test, MODEL_GB.predict(DATA_MODEL_GB.X_test))
                return(MATTHEWS_CORRCOEF)
    

        # Search for best hyperparameters
        study = optuna.create_study(direction='maximize')
        study.optimize(objective_GB, n_trials=Global_Parameters.GB_MODEL_TRIAL, catch=(ValueError,))
        
        # Saving and using best hyperparameters
        BEST_PARAMS_GB = np.zeros([1], dtype=object)
        BEST_PARAMS_GB[0] = study.best_params
        DATA_MODEL_GB.learning_rate = float(BEST_PARAMS_GB[0].get("learning_rate"))
        DATA_MODEL_GB.Nb_Tree = int(BEST_PARAMS_GB[0].get("Nb_Tree"))
        DATA_MODEL_GB.min_samples_leaf = int(BEST_PARAMS_GB[0].get("min_samples_leaf"))
        DATA_MODEL_GB.min_samples_split = int(BEST_PARAMS_GB[0].get("min_samples_split"))
        DATA_MODEL_GB.min_weight_fraction_leaf = float(BEST_PARAMS_GB[0].get("min_weight_fraction_leaf"))
        DATA_MODEL_GB.max_depth = int(BEST_PARAMS_GB[0].get("max_depth"))
    
    # Creating and fitting gradient boosting model
    DATA_MODEL_GB.gradient_boosting_modellisation(Global_Parameters.N_SPLIT, Global_Parameters.REGRESSION)
    
    #
    # Result analysis
    
    # Classification
    if Global_Parameters.REGRESSION == False:
        DATA_MODEL_GB.result_plot_classification()
        DATA_MODEL_GB.plot_training_validation_error()
        DATA_MODEL_GB.result_report_classification_calculation()
        DATA_MODEL_GB.result_report_classification_print()
        DATA_MODEL_GB.result_report_classification_plot()
    
    # Regression
    else:
        DATA_MODEL_GB.result_plot_regression(Global_Data.TRAIN_DATAFRAME, Global_Parameters.NAME_DATA_PREDICT)
        DATA_MODEL_GB.result_report_regression_calculation()
        DATA_MODEL_GB.result_report_regression_print()
        DATA_MODEL_GB.result_report_regression_plot()
        
        DATA_MODEL_GB.extract_max_diff_regression()
    
    # Exit
    return DATA_MODEL_GB