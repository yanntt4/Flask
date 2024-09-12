# -*- coding: utf-8 -*-
"""
Created on Thu Aug 29 16:01:52 2024

@author: ythiriet
"""

# Global library import
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plot
import random

from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.model_selection import cross_val_score
from sklearn.inspection import permutation_importance

from sklearn.metrics import matthews_corrcoef

import shap
import optuna


# Personal library
sys.path.append("./")
from Data_modelling import Data_modelling


# Class definition
class Data_modelling_random_forest(Data_modelling):
    def __init__(self):
        super(Data_modelling_random_forest, self).__init__()

        self.Nb_Tree = 146
        self.min_samples_leaf = 16
        self.min_samples_split = 7
        self.min_weight_fraction_leaf = 0.00007276912136637689
        self.max_depth = 33
        
        self.Nb_Tree = 140
        self.min_samples_leaf = 5
        self.min_samples_split = 22
        self.min_weight_fraction_leaf = 0.00021061239694571002
        self.max_depth = 32

        self.START_POINT_SHAP = 0
        self.END_POINT_SHAP_SMALL = 20
        self.END_POINT_SHAP_LONG = 100
        self.SHAP_EXPLAINER = 0


    # Fitting random forest with data
    def random_forest_modellisation(self, K_FOLDS, REGRESSION):

        # Setting the model with parameters
        if REGRESSION:
            self.MODEL = RandomForestRegressor(
                n_estimators=self.Nb_Tree,
                min_samples_leaf=self.min_samples_leaf,
                min_samples_split=self.min_samples_split,
                min_weight_fraction_leaf=self.min_weight_fraction_leaf,
                max_depth=self.max_depth,
                verbose=2,
                random_state=0)
        
        else:
            self.MODEL = RandomForestClassifier(
                n_estimators=self.Nb_Tree,
                min_samples_leaf=self.min_samples_leaf,
                min_samples_split=self.min_samples_split,
                min_weight_fraction_leaf=self.min_weight_fraction_leaf,
                max_depth=self.max_depth,
                verbose=2,
                random_state=0)

        # Cross validation
        self.score = cross_val_score(self.MODEL, self.X_train, self.Y_train, cv=K_FOLDS)
        self.MODEL.fit(self.X_train, self.Y_train)

        # Predicting results
        self.Y_predict = self.MODEL.predict(self.X_test)
        Y_test = np.squeeze(self.Y_test.to_numpy())

        # Percentage calculation for correct prediction
        if REGRESSION:
            self.AVERAGE_DIFFERENCE = np.mean(abs(self.Y_predict - Y_test))
            print(f"\n Moyenne des différences : {round(self.AVERAGE_DIFFERENCE,2)} €")
            self.PERCENTAGE_AVERAGE_DIFFERENCE = 100*self.AVERAGE_DIFFERENCE / np.mean(self.Y_test)
            print(f"\n Pourcentage de différence : {round(self.PERCENTAGE_AVERAGE_DIFFERENCE,2)} %")
            
        else:
            self.Y_PREDICT_PROBA = self.MODEL.predict_proba(self.X_test)
            self.NB_CORRECT_PREDICTION = np.count_nonzero(
                Y_test.astype(int) - self.Y_predict.astype(int))
            self.PERCENTAGE_CORRECT_PREDICTION = (1 -
                self.NB_CORRECT_PREDICTION / Y_test.shape[0])
            print(f"\n Pourcentage de predictions correctes {self.MODEL_NAME} : {100*round(self.PERCENTAGE_CORRECT_PREDICTION,5)} %")


    # Plotting feature importance
    def feature_importance_plot(self):

        # Feature Importance
        RF_Feature_Importance = pd.DataFrame(
            {'Variable': self.X_train.columns,
              'Importance': self.MODEL.feature_importances_}).sort_values(
                  'Importance', ascending=False)

        fig, ax = plot.subplots()
        ax.barh(RF_Feature_Importance.Variable, RF_Feature_Importance.Importance)
        plot.grid()
        plot.suptitle("Feature Importance for Random Forest Model")


    # Plotting permutation importance
    def permutation_importance_plot(self, TEST_DATAFRAME):

          # Permutation Importance
        PERMUTATION_IMPORTANCE_TRAIN = permutation_importance(
            self.MODEL, self.X_train, self.Y_train, n_repeats=10, random_state=0, n_jobs=2)
        PERMUTATION_IMPORTANCE_TEST = permutation_importance(
            self.MODEL, self.X_test, self.Y_test, n_repeats=10, random_state=0, n_jobs=2)

        # Init
        fig, ax = plot.subplots(2)
        max_importances = 0

        # Loop for Train/Test Data
        for i, permutation_importance_plot in enumerate(
                [PERMUTATION_IMPORTANCE_TRAIN, PERMUTATION_IMPORTANCE_TEST]):

            # Calculating permutaion importance
            sorted_importances_idx = permutation_importance_plot.importances_mean.argsort()
            importances = pd.DataFrame(
                permutation_importance_plot.importances[sorted_importances_idx].T,
                columns=self.X_test.columns[sorted_importances_idx],)
            max_importances = max([max_importances,importances.max().max()])

            # Plotting results
            ax[i].boxplot(importances, vert=False)
            ax[i].set_title("Permutation Importances")
            ax[i].axvline(x=0, color="k", linestyle="--")
            ax[i].set_xlabel("Decrease in accuracy score")
            ax[i].set_xlim([-0.01,max_importances + 0.1])
            ax[i].set_yticks(np.linspace(1,importances.shape[1],importances.shape[1]))
            ax[i].set_yticklabels(importances.columns)
            ax[i].figure.tight_layout()


    # Plotting shap values for a single point
    def shap_value_analysis_single_point(self):

        # Init
        X_TEST_RESET_INDEX = self.X_test.reset_index()
        NB_ANALYSED = X_TEST_RESET_INDEX.iloc[random.randint(0,X_TEST_RESET_INDEX.shape[0])].name
        INDEX_ANALYSED = X_TEST_RESET_INDEX.iloc[NB_ANALYSED,:]["index"]

        # Create object that can calculate shap values
        self.SHAP_EXPLAINER = shap.TreeExplainer(self.MODEL)
        # Expected value is the mean result

        # Calculate Shap values for one point
        self.MODEL.predict(np.array(self.X_test.iloc[NB_ANALYSED,:]).reshape(-1,1).T)
        SHAP_VALUES_ONE = self.SHAP_EXPLAINER.shap_values(self.X_test.iloc[NB_ANALYSED,:])

        # Plotting results
        shap.initjs()
        shap.force_plot(
            self.SHAP_EXPLAINER.expected_value[0],
            SHAP_VALUES_ONE[:,1],
            self.X_test.iloc[NB_ANALYSED,:].index,
            matplotlib=True)
        plot.suptitle(f"Prediction attendue : {self.Y_test[INDEX_ANALYSED]}")


    # Plotting shap values for 10 points
    def shap_value_analysis_multiple_point(self):

        # Calculate Shap values for 10 points
        WRONG_PRED = self.MODEL.predict(self.X_test) != np.array(self.Y_test)
        SHAP_VALUES_TEN = self.SHAP_EXPLAINER.shap_values(
            self.X_test.iloc[self.START_POINT_SHAP:self.END_POINT_SHAP_SMALL,:])

        # Plotting results
        plot.figure()
        shap.decision_plot(
            self.SHAP_EXPLAINER.expected_value[0],
            SHAP_VALUES_TEN[:,:,0],
            self.X_test.iloc[self.START_POINT_SHAP:self.END_POINT_SHAP_SMALL,:],
            feature_names = np.array(self.X_test.columns),
            link="logit",
            highlight=WRONG_PRED[self.START_POINT_SHAP:self.END_POINT_SHAP_SMALL])
        

    # Plotting shap values for 100 points
    def shap_value_analysis_multiple_massive_point(self):

        # Calculate Shap values for 100 points
        SHAP_VALUES_HUNDRED = self.SHAP_EXPLAINER.shap_values(
            self.X_test.iloc[self.START_POINT_SHAP:self.END_POINT_SHAP_LONG,:])

        # Plotting results
        shap.dependence_plot(2, SHAP_VALUES_HUNDRED[:,:,0], 
                              self.X_test.iloc[self.START_POINT_SHAP:self.END_POINT_SHAP_LONG,:])



# Function to create/optimize Random Forest model
def random_forest(Data_Model, Global_Parameters, Global_Data):
    DATA_MODEL_RF = Data_modelling_random_forest()
    
    # Using split create previously
    DATA_MODEL_RF.X_train = Data_Model.X_train
    DATA_MODEL_RF.Y_train = Data_Model.Y_train
    DATA_MODEL_RF.X_test = Data_Model.X_test
    DATA_MODEL_RF.Y_test = Data_Model.Y_test
    DATA_MODEL_RF.MODEL_NAME = "Random Forest"
    
    #
    # Building a Random Forest Model with adjusted parameters
    
    # Regression
    if Global_Parameters.REGRESSION:
        def build_model_RF(Nb_Tree=1, min_samples_leaf=2, min_samples_split=10,
                           max_depth=2, min_weight_fraction_leaf=0.5):
    
            MODEL_RF = RandomForestRegressor(
                n_estimators=Nb_Tree,
                min_samples_leaf=min_samples_leaf,
                min_samples_split=min_samples_split,
                random_state=0,
                max_depth=max_depth,
                min_weight_fraction_leaf=min_weight_fraction_leaf,)
    
            return MODEL_RF
    
    # Classification
    else:
        def build_model_RF(Nb_Tree=1, min_samples_leaf=2, min_samples_split=10,
                           max_depth=2, min_weight_fraction_leaf=0.5):
    
            MODEL_RF = RandomForestClassifier(
                n_estimators=Nb_Tree,
                min_samples_leaf=min_samples_leaf,
                min_samples_split=min_samples_split,
                random_state=0,
                max_depth=max_depth,
                min_weight_fraction_leaf=min_weight_fraction_leaf,)
    
            return MODEL_RF
    
    
    # Searching for Optimized Hyperparameters
    if Global_Parameters.RF_MODEL_OPTI:
        
        # Regression
        if Global_Parameters.REGRESSION:
    
            # Building function to minimize/maximize score
            def objective_RF(trial):
                params = {'Nb_Tree': trial.suggest_int('Nb_Tree', 10, 250),
                          'min_samples_leaf': trial.suggest_int('min_samples_leaf', 2, 20),
                          'min_samples_split': trial.suggest_int('min_samples_split', 2, 50),
                          'max_depth': trial.suggest_int('max_depth', 1, 50),
                          'min_weight_fraction_leaf': trial.suggest_float('min_weight_fraction_leaf', 0, 0.5)}
        
                MODEL_RF = build_model_RF(**params)
                MODEL_RF.fit(DATA_MODEL_RF.X_train, DATA_MODEL_RF.Y_train)
                scores = cross_val_score(MODEL_RF, DATA_MODEL_RF.X_test, DATA_MODEL_RF.Y_test, cv=Global_Parameters.k_folds)
                prediction_score = MODEL_RF.score(DATA_MODEL_RF.X_test, DATA_MODEL_RF.Y_test)
                return (4*np.mean(scores) + prediction_score)/5

        # Classification
        else:
            
            # Building function to minimize/maximize score
            def objective_RF(trial):
                params = {'Nb_Tree': trial.suggest_int('Nb_Tree', 10, 250),
                          'min_samples_leaf': trial.suggest_int('min_samples_leaf', 2, 20),
                          'min_samples_split': trial.suggest_int('min_samples_split', 2, 50),
                          'max_depth': trial.suggest_int('max_depth', 1, 50),
                          'min_weight_fraction_leaf': trial.suggest_float('min_weight_fraction_leaf', 0, 0.5)}
        
                MODEL_RF = build_model_RF(**params)
                MODEL_RF.fit(DATA_MODEL_RF.X_train, DATA_MODEL_RF.Y_train)
                
                MATTHEWS_CORRCOEF = matthews_corrcoef(DATA_MODEL_RF.Y_test, MODEL_RF.predict(DATA_MODEL_RF.X_test))
                return(MATTHEWS_CORRCOEF)
            
            
    
        # Search for best parameters
        study = optuna.create_study(direction='maximize')
        study.optimize(objective_RF, n_trials=Global_Parameters.RF_MODEL_TRIAL,
                       catch=(ValueError,))
        
        # Saving and using best hyperparameters
        BEST_PARAMS_RF = np.zeros([1], dtype=object)
        BEST_PARAMS_RF[0] = study.best_params
        DATA_MODEL_RF.Nb_Tree = int(BEST_PARAMS_RF[0].get("Nb_Tree"))
        DATA_MODEL_RF.min_samples_leaf = int(BEST_PARAMS_RF[0].get("min_samples_leaf"))
        DATA_MODEL_RF.min_samples_split = int(BEST_PARAMS_RF[0].get("min_samples_split"))
        DATA_MODEL_RF.min_weight_fraction_leaf = float(BEST_PARAMS_RF[0].get("min_weight_fraction_leaf"))
        DATA_MODEL_RF.max_depth = int(BEST_PARAMS_RF[0].get("max_depth"))
    
    # Creating and fitting random forest model
    DATA_MODEL_RF.random_forest_modellisation(Global_Parameters.k_folds, Global_Parameters.REGRESSION)
    
    #
    # Result analysis
    DATA_MODEL_RF.feature_importance_plot()
    DATA_MODEL_RF.permutation_importance_plot(Global_Data.TEST_DATAFRAME)
    
    # Classification
    if Global_Parameters.REGRESSION == False:
        DATA_MODEL_RF.result_plot_classification()
        DATA_MODEL_RF.result_report_classification_calculation()
        DATA_MODEL_RF.result_report_classification_print()
        DATA_MODEL_RF.result_report_classification_plot()
        
        DATA_MODEL_RF.shap_value_analysis_single_point()
        DATA_MODEL_RF.shap_value_analysis_multiple_point()
        DATA_MODEL_RF.shap_value_analysis_multiple_massive_point()
        
    # Regression
    else:
        DATA_MODEL_RF.result_plot_regression(Global_Data.TRAIN_DATAFRAME, Global_Parameters.NAME_DATA_PREDICT)
        DATA_MODEL_RF.result_report_regression_calculation()
        DATA_MODEL_RF.result_report_regression_print()
        DATA_MODEL_RF.result_report_regression_plot()
        
        DATA_MODEL_RF.extract_max_diff_regression()
    
    # Exit
    return DATA_MODEL_RF