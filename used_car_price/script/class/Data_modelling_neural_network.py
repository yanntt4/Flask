# -*- coding: utf-8 -*-
"""
Created on Thu Aug 29 16:28:24 2024

@author: ythiriet
"""

# Global library import
import sys
import numpy as np
import pandas as pd
import matplotlib.pyplot as plot

import tensorflow as tf

import optuna
from sklearn.metrics import accuracy_score, mean_squared_error

# Personal library
sys.path.append("./")
from Data_modelling import Data_modelling


# Class definition
class Data_modelling_neural_network(Data_modelling):
    def __init__(self):
        super(Data_modelling_neural_network, self).__init__()

        self.n_hidden = 3
        self.n_neurons = 68
        self.n_trials = 20
        self.History = 0

        self.monitor ="val_accuracy"
        self.min_delta = 0.002
        # self.patience = 25
        self.patience = 5
    
    
    def neural_network_modellisation(self, N_SPLIT, REGRESSION, UNIQUE_PREDICT_VALUES):
        
        tf.random.set_seed(0)

        if REGRESSION:
            
            # Neural Network model
            self.MODEL = tf.keras.models.Sequential()
            self.MODEL.add(tf.keras.layers.InputLayer(input_shape=(self.X_train.shape[1],)))
            for layers in range(self.n_hidden):
                self.MODEL.add(tf.keras.layers.Dense(self.n_neurons, activation = "relu"))
            self.MODEL.add(tf.keras.layers.Dense(1))
    
            # Optimizer used
            OPTIMIZER = tf.keras.optimizers.Nadam()
    
            self.MODEL.compile(loss = "mse",
                               optimizer = OPTIMIZER,
                               metrics = ["accuracy"])
        
        else:

            # Neural Network model
            self.MODEL = tf.keras.models.Sequential()
            self.MODEL.add(tf.keras.layers.InputLayer(input_shape=(self.X_train.shape[1],)))
            for layers in range(self.n_hidden):
                self.MODEL.add(tf.keras.layers.Dense(self.n_neurons, activation = "relu"))
            self.MODEL.add(tf.keras.layers.Dense(UNIQUE_PREDICT_VALUES.shape[0], activation = "softmax"))
    
            # Optimizer used
            OPTIMIZER = tf.keras.optimizers.Nadam()
    
            self.MODEL.compile(
                loss = "sparse_categorical_crossentropy", optimizer = OPTIMIZER, metrics = ["accuracy"])


        # Early stopping init
        callback = tf.keras.callbacks.EarlyStopping(monitor = self.monitor,
                                                    min_delta = self.min_delta,
                                                    patience = self.patience,
                                                    verbose = 0,
                                                    restore_best_weights=True)
        self.History = self.MODEL.fit(np.asarray(self.X_train).astype("float32"),
                                      np.asarray(self.Y_train).astype("float32"),
                                      epochs = 500,
                                      validation_split = 0.01,
                                      initial_epoch = 0,
                                      callbacks=[callback])
        
        # Plot learning evolution for Neural Network
        pd.DataFrame(self.History.history).plot(figsize = (8,5))
        plot.grid(True)
        plot.title("Learning Evolution for Neural Network")

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
            self.Y_PREDICT_PROBA = self.Y_predict
            self.Y_predict = np.argmax(self.Y_PREDICT_PROBA,axis=1)
            self.NB_CORRECT_PREDICTION = np.count_nonzero(
                Y_test.astype(int) - self.Y_predict.astype(int))
            self.PERCENTAGE_CORRECT_PREDICTION = (1 -
                self.NB_CORRECT_PREDICTION / Y_test.shape[0])
            print(f"\n Pourcentage de predictions correctes {self.MODEL_NAME} : {100*round(self.PERCENTAGE_CORRECT_PREDICTION,5)} %")


# Function to create/optimize Neural Network model
def neural_network(Data_Model, Global_Parameters, Global_Data):
    DATA_MODEL_NN = Data_modelling_neural_network()
    
    # Using split create previously
    DATA_MODEL_NN.X_train = Data_Model.X_train
    DATA_MODEL_NN.Y_train = Data_Model.Y_train
    DATA_MODEL_NN.X_test = Data_Model.X_test
    DATA_MODEL_NN.Y_test = Data_Model.Y_test
    DATA_MODEL_NN.MODEL_NAME = "Neural Network"
    
    #
    # Building a neural network modem with adjusted parameters
    
    # Regression
    if Global_Parameters.REGRESSION:
        
        def build_model_NN(n_hidden = 1, n_neurons = 100, input_shape = (Data_Model.X_train.shape[1],)):
    
            # Neural Network model
            MODEL = tf.keras.models.Sequential()
            MODEL.add(tf.keras.layers.InputLayer(input_shape=input_shape))
            for layers in range(n_hidden):
                MODEL.add(tf.keras.layers.Dense(n_neurons, activation = "relu"))
            MODEL.add(tf.keras.layers.Dense(1))
    
            # Optimizer used
            OPTIMIZER = tf.keras.optimizers.Nadam()
    
            MODEL.compile(loss = "mse",
                          optimizer = OPTIMIZER,
                          metrics = ["accuracy"])
    
            return MODEL
    
    # Classification
    else:

        def build_model_NN(n_hidden = 1, n_neurons = 100, input_shape = (Data_Model.X_train.shape[1],)):
    
            # Neural Network model
            MODEL = tf.keras.models.Sequential()
            MODEL.add(tf.keras.layers.InputLayer(input_shape=input_shape))
            for layers in range(n_hidden):
                MODEL.add(tf.keras.layers.Dense(n_neurons, activation = "relu"))
            MODEL.add(tf.keras.layers.Dense(Global_Data.UNIQUE_PREDICT_VALUE.shape[0], activation = "softmax"))
    
            # Optimizer used
            OPTIMIZER = tf.keras.optimizers.Nadam()
    
            MODEL.compile(loss = "sparse_categorical_crossentropy",
                          optimizer = OPTIMIZER,
                          metrics = ["accuracy"])
    
            return MODEL


    if Global_Parameters.NN_MODEL_OPTI:
        
        # Regression
        if Global_Parameters.REGRESSION:
            
            # Building function to minimize/maximise
            def objective_NN(trial):
                params = {
                    'n_hidden': trial.suggest_int('n_hidden', 2, 4),
                    'n_neurons': trial.suggest_int('n_neurons', 10, 100)}
    
                model = build_model_NN(**params)
    
                model.fit(np.asarray(DATA_MODEL_NN.X_train).astype("float32"),
                          np.asarray(DATA_MODEL_NN.Y_train).astype("float32"),
                          epochs = 50,
                          validation_split = 0.01,
                          initial_epoch = 0)
    
                Preds_NN = model.predict(np.asarray(DATA_MODEL_NN.X_test).astype("float32"))
                Score_NN = mean_squared_error(Data_Model.Y_test, Preds_NN)
    
                return Score_NN
            
        # Classification
        else:

            # Building function to minimize/maximise
            def objective_NN(trial):
                params = {
                    'n_hidden': trial.suggest_int('n_hidden', 2, 4),
                    'n_neurons': trial.suggest_int('n_neurons', 10, 100)}
    
                model = build_model_NN(**params)
    
                model.fit(np.asarray(DATA_MODEL_NN.X_train).astype("float32"),
                          np.asarray(DATA_MODEL_NN.Y_train).astype("float32"),
                          epochs = 50,
                          validation_split = 0.01,
                          initial_epoch = 0)
    
                Preds_NN_proba = model.predict(np.asarray(DATA_MODEL_NN.X_test).astype("float32"))
                Preds_NN = np.zeros([Preds_NN_proba.shape[0]], dtype = int)
    
                # Turning probability prediction into prediction
                for i in range(Preds_NN_proba.shape[0]):
                    Preds_NN[i] = np.where(Preds_NN_proba[i,:] == np.amax(Preds_NN_proba[i,:]))[0][0]
    
                Score_NN = accuracy_score(Data_Model.Y_test, Preds_NN)
    
                return Score_NN


        # Search for best hyperparameters
        if Global_Parameters.REGRESSION:
            study = optuna.create_study(direction='minimize')
        else:
            study = optuna.create_study(direction='maximize')
        study.optimize(objective_NN,
                       n_trials=Global_Parameters.NN_MODEL_TRIAL,
                       catch=(ValueError,))
        
        # Saving and using best hyperparameters
        BEST_PARAMS_NN = np.zeros([1], dtype = object)
        BEST_PARAMS_NN[0] = study.best_params
        DATA_MODEL_NN.n_hidden = int(BEST_PARAMS_NN[0].get("n_hidden"))
        DATA_MODEL_NN.n_neurons = int(BEST_PARAMS_NN[0].get("n_neurons"))
        
    # Creating and fitting xgboosting model
    DATA_MODEL_NN.neural_network_modellisation(
        Global_Parameters.N_SPLIT, Global_Parameters.REGRESSION, Global_Data.UNIQUE_PREDICT_VALUE)
    
  
    #
    # Result analysis
    
    # Classification
    if Global_Parameters.REGRESSION == False:
        DATA_MODEL_NN.result_plot_classification()
        DATA_MODEL_NN.result_report_classification_calculation()
        DATA_MODEL_NN.result_report_classification_print()
        DATA_MODEL_NN.result_report_classification_plot()

    # Regression
    else:
        DATA_MODEL_NN.result_plot_regression(Global_Data.TRAIN_DATAFRAME, Global_Parameters.NAME_DATA_PREDICT)
        DATA_MODEL_NN.result_report_regression_calculation()
        DATA_MODEL_NN.result_report_regression_print()
        DATA_MODEL_NN.result_report_regression_plot()
        
        DATA_MODEL_NN.extract_max_diff_regression()
    
    # Exit
    return DATA_MODEL_NN