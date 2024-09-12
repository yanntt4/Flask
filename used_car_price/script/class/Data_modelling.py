# -*- coding: utf-8 -*-
"""
Created on Thu Aug 29 15:42:08 2024

@author: ythiriet
"""


# Global library import
import numpy as np
import pandas as pd
import matplotlib.pyplot as plot

from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, mean_squared_error, mean_absolute_error, r2_score
from sklearn.metrics import balanced_accuracy_score, precision_score, recall_score, f1_score
from sklearn.metrics import confusion_matrix, classification_report, matthews_corrcoef
from sklearn.metrics import explained_variance_score, max_error, root_mean_squared_error
from sklearn.metrics import mean_squared_log_error, root_mean_squared_log_error
from sklearn.metrics import median_absolute_error, mean_poisson_deviance
from sklearn.metrics import mean_gamma_deviance, mean_absolute_percentage_error
from sklearn.metrics import d2_absolute_error_score

from imblearn.over_sampling import SMOTEN
from imblearn.metrics import classification_report_imbalanced


# Class definition
class Data_modelling():
    def __init__(self):
        self.X_train = []
        self.Y_train = []
        self.X_test = []
        self.Y_test = []
        self.Y_predict = []
        self.Y_PREDICT_PROBA = []
        self.K_predict = []
        self.K_predict_proba = []

        self.MODEL = ""
        self.MODEL_NAME = ""
        self.Y_predict = []
        self.NB_CORRECT_PREDICTION = 0
        self.PERCENTAGE_CORRECT_PREDICTION = 0
        self.score = 0
        self.BEST_PARAMS = np.zeros([1], dtype = object)


    # Splitting data into train and test dataframe to avoid overfitting
    def splitting_data(self, TRAIN_DATAFRAME, GENERIC_NAME_DATA_PREDICT, MULTI_CLASSIFICATION, REGRESSION):

        # Split data creation
        self.X_train, self.X_test, self.Y_train, self.Y_test = train_test_split(
            TRAIN_DATAFRAME.drop([GENERIC_NAME_DATA_PREDICT], axis=1),
            TRAIN_DATAFRAME.loc[:, [GENERIC_NAME_DATA_PREDICT]].iloc[:,0],
            test_size=0.2,
            random_state=0)

        # Turning Y_train and Y_test into boolean if needed
        if MULTI_CLASSIFICATION == False and REGRESSION == False:
            self.Y_train = self.Y_train.astype(bool)
            self.Y_test = self.Y_test.astype(bool)

    
    # Smoten data to complete missing class for classification
    def smoten_sampling(self):

        sampler = SMOTEN(random_state = 0)
        self.X_train, self.Y_train = sampler.fit_resample(self.X_train, self.Y_train)
    
    
    # Searching for max difference between prediction and reality 
    def extract_max_diff_regression(self):
        
        # Finding data with highest difference between real and prediction
        X_TEST_EXTREMA_ANALYSIS = self.X_test.copy()
        X_TEST_EXTREMA_ANALYSIS["Real"] = self.Y_test
        X_TEST_EXTREMA_ANALYSIS["Predicted"] = self.Y_predict
        X_TEST_EXTREMA_ANALYSIS["Difference"] = abs(self.Y_test - self.Y_predict)
        self.X_TEST_EXTREMA_INDEX = X_TEST_EXTREMA_ANALYSIS.sort_values(by = ["Difference"]).nlargest(10, "Difference").index
        
        # Data standardisation using max value
        self.X_TEST_STANDARD = self.X_test.copy()
        
        for COLUMN in self.X_test.columns:
            self.X_TEST_STANDARD[COLUMN] = self.X_TEST_STANDARD[COLUMN]/self.X_TEST_STANDARD[COLUMN].max()
        
        # Finiding closest points using euclidian distance
        X_TEST_STANDARD_EXTREMA = self.X_TEST_STANDARD[self.X_TEST_STANDARD.index == self.X_TEST_EXTREMA_INDEX[0]]
        
        EUCLIDIAN_DISTANCE_CALCULATION_ARRAY = np.zeros([self.X_TEST_STANDARD.shape[0]], dtype = float)
        X_TEST_STANDARD_ARRAY = np.array(self.X_TEST_STANDARD)
        X_TEST_STANDARD_EXTREMA_ARRAY = np.array(X_TEST_STANDARD_EXTREMA)
        
        for i in range(X_TEST_STANDARD_ARRAY.shape[0]):
            for j in range(X_TEST_STANDARD_ARRAY.shape[1]):
                EUCLIDIAN_DISTANCE_CALCULATION_ARRAY[i] += (
                    (X_TEST_STANDARD_ARRAY[i,j] - X_TEST_STANDARD_EXTREMA_ARRAY[0,j])*(X_TEST_STANDARD_ARRAY[i,j] - X_TEST_STANDARD_EXTREMA_ARRAY[0,j]))
        
        self.X_TEST_STANDARD["Euclidian Distance"] = EUCLIDIAN_DISTANCE_CALCULATION_ARRAY
        self.X_TEST_CLOSEST_EXTREMA = self.X_test.loc[np.array(self.X_TEST_STANDARD.sort_values(by = ["Euclidian Distance"]).nlargest(6, "Euclidian Distance").index)]
        self.X_TEST_CLOSEST_EXTREMA["SalePrice"] = self.Y_test.loc[np.array(self.X_TEST_STANDARD.sort_values(by = ["Euclidian Distance"]).nlargest(6, "Euclidian Distance").index)]
        

    # Plotting result with classification
    def result_plot_classification(self):

        # Plotting results
        X_plot = np.linspace(1, self.Y_test.shape[0], self.Y_test.shape[0])

        fig, ax = plot.subplots(2)
        ax[0].scatter(X_plot, self.Y_predict, color="green")
        ax[0].scatter(X_plot, self.Y_test, color="orange")
        ax[0].legend([f"Prediction from {self.MODEL_NAME} Model", "Real Results"])
        ax[1].scatter(X_plot, np.sort(abs(self.Y_test.astype(int) - self.Y_predict.astype(int))))
        ax[1].set_title("Difference between predict and real results")
        ax[1].legend([f"Score for {self.MODEL_NAME} Model : {round(self.PERCENTAGE_CORRECT_PREDICTION,2)}"])
        plot.grid()
    
    
    # Plitting result with regression
    def result_plot_regression(self, TRAIN_DATAFRAME, NAME_DATA_PREDICT):
        
        # Preparation
        QUANTILE = TRAIN_DATAFRAME[NAME_DATA_PREDICT].quantile([0.25,0.5,0.75])
        self.Y_predict = np.array(self.Y_predict)
        if self.Y_predict.ndim == 2:
            self.Y_predict = self.Y_predict[:,0]

        # Plotting results
        X_plot = np.linspace(1, self.Y_test.shape[0], self.Y_test.shape[0])

        fig, ax = plot.subplots(3)
        ax[0].scatter(X_plot, self.Y_predict, color="green")
        ax[0].scatter(X_plot, self.Y_test, color="orange")
        ax[0].legend([f"Prediction from {self.MODEL_NAME} Model", "Real Results"])
        ax[1].scatter(X_plot, np.sort(abs(self.Y_test - self.Y_predict)))
        ax[1].set_title("Difference between predict and real results")
        ax[1].legend([f"Score for {self.MODEL_NAME} Model : {round(self.AVERAGE_DIFFERENCE,2)}"])
        plot.grid()
        ax[2].boxplot(abs(self.Y_test - self.Y_predict), vert = False)
        
        
        self.PREDICT_TEST_DATAFRAME = pd.DataFrame(self.Y_predict, columns = ["Prediction"])
        self.PREDICT_TEST_DATAFRAME["Reel"] = self.Y_test.reset_index(drop = True)
        self.PREDICT_TEST_DATAFRAME["Diff predict reel"] = (self.PREDICT_TEST_DATAFRAME["Prediction"] - self.PREDICT_TEST_DATAFRAME["Reel"])
        
        fig, ax = plot.subplots(2,2)
        ax[0,0].boxplot(self.PREDICT_TEST_DATAFRAME["Diff predict reel"].loc[self.PREDICT_TEST_DATAFRAME["Prediction"] < QUANTILE[0.25]],
                        vert = False)
        ax[0,0].set_title(f"< 1st quantile // Average : {round(self.PREDICT_TEST_DATAFRAME['Diff predict reel'].loc[self.PREDICT_TEST_DATAFRAME['Prediction'] < QUANTILE[0.25]].mean(),1)}")
        ax[0,1].boxplot(self.PREDICT_TEST_DATAFRAME["Diff predict reel"].loc[(self.PREDICT_TEST_DATAFRAME["Prediction"] < QUANTILE[0.5]) &
                                                                             (self.PREDICT_TEST_DATAFRAME["Prediction"] > QUANTILE[0.25])],
                        vert = False)
        ax[0,1].set_title(f"1st quantile to median // Average : {round(self.PREDICT_TEST_DATAFRAME['Diff predict reel'].loc[(self.PREDICT_TEST_DATAFRAME['Prediction'] < QUANTILE[0.5]) & (self.PREDICT_TEST_DATAFRAME['Prediction'] > QUANTILE[0.25])].mean(),1)}")
        ax[1,0].boxplot(self.PREDICT_TEST_DATAFRAME["Diff predict reel"].loc[(self.PREDICT_TEST_DATAFRAME["Prediction"] < QUANTILE[0.75]) &
                                                                             (self.PREDICT_TEST_DATAFRAME["Prediction"] > QUANTILE[0.5])],
                        vert = False)
        ax[1,0].set_title(f"median to 3rd quantile // Average : {round(self.PREDICT_TEST_DATAFRAME['Diff predict reel'].loc[(self.PREDICT_TEST_DATAFRAME['Prediction'] < QUANTILE[0.75]) & (self.PREDICT_TEST_DATAFRAME['Prediction'] > QUANTILE[0.5])].mean(),1)}")
        ax[1,1].boxplot(self.PREDICT_TEST_DATAFRAME["Diff predict reel"].loc[self.PREDICT_TEST_DATAFRAME["Prediction"] > QUANTILE[0.75]],
                        vert = False)
        ax[1,1].set_title(f"> 3rd quantile // Average : {round(self.PREDICT_TEST_DATAFRAME['Diff predict reel'].loc[self.PREDICT_TEST_DATAFRAME['Prediction'] > QUANTILE[0.75]].mean(),1)}")
        plot.suptitle("Difference between Prediction and Reality between quantile")


    # KPI calculation with classification
    def result_report_classification_calculation(self):
        self.CONFUSION_MATRIX = confusion_matrix(self.Y_test, self.Y_predict)
        self.ACCURACY = accuracy_score(self.Y_test, self.Y_predict)
        self.BALANCED_ACCURACY = balanced_accuracy_score(self.Y_test, self.Y_predict)
        
        # Precision
        self.MICRO_PRECISION = precision_score(self.Y_test, self.Y_predict, average='micro')
        self.MACRO_PRECISION = precision_score(self.Y_test, self.Y_predict, average='macro')
        self.WEIGHTED_PRECISION = precision_score(self.Y_test, self.Y_predict, average='weighted')
        
        # Recall
        self.MICRO_RECALL = recall_score(self.Y_test, self.Y_predict, average='micro')
        self.MACRO_RECALL = recall_score(self.Y_test, self.Y_predict, average='macro')
        self.WEIGHTED_RECALL = recall_score(self.Y_test, self.Y_predict, average='weighted')
        
        # F1-score
        self.MICRO_F1_SCORE = f1_score(self.Y_test, self.Y_predict, average='micro')
        self.MACRO_F1_SCORE = f1_score(self.Y_test, self.Y_predict, average='macro')
        self.WEIGHTED_F1_SCORE = f1_score(self.Y_test, self.Y_predict, average='weighted')
        
        # Matthews correlation coefficient (MCC)
        self.MCC = matthews_corrcoef(self.Y_test, self.Y_predict)


    # KPI print with classification
    def result_report_classification_print(self):

        print('\n------------------ Confusion Matrix -----------------\n')
        print(self.CONFUSION_MATRIX)

        print('\n-------------------- Key Metrics --------------------')
        print('\nAccuracy: {:.3f}'.format(self.ACCURACY))
        print('Balanced Accuracy: {:.3f}\n'.format(self.BALANCED_ACCURACY))

        print('Micro Precision: {:.3f}'.format(self.MICRO_PRECISION))
        print('Micro Recall: {:.3f}'.format(self.MICRO_RECALL))
        print('Micro F1-score: {:.3f}\n'.format(self.MICRO_F1_SCORE))

        print('Macro Precision: {:.3f}'.format(self.MACRO_PRECISION))
        print('Macro Recall: {:.3f}'.format(self.MACRO_RECALL))
        print('Macro F1-score: {:.3f}\n'.format(self.MACRO_F1_SCORE))

        print('Weighted Precision: {:.3f}'.format(self.WEIGHTED_PRECISION))
        print('Weighted Recall: {:.3f}'.format(self.WEIGHTED_RECALL))
        print('Weighted F1-score: {:.3f}\n'.format(self.WEIGHTED_F1_SCORE))
        
        print('Matthews correlation coefficient: {:.3f}\n'.format(self.MCC))

        print('\n--------------- Classification Report ---------------\n')
        print(classification_report(self.Y_test, self.Y_predict))

        print('\n--------------- Imbalanced Report ---------------\n')
        print(classification_report_imbalanced(self.Y_test, self.Y_predict))


    # KPI plot with classification 
    def result_report_classification_plot(self):

        plot.figure(figsize = (10,8))
        plot.ylim(1,40)
        plot.text(0.02,39,'------------------ Confusion Matrix -----------------')
        plot.text(0.02,29, self.CONFUSION_MATRIX)
        plot.text(0.4,28,'-------------------- Key Metrics --------------------')
        plot.text(0.4,26,'Accuracy: {:.3f}'.format(self.ACCURACY))
        plot.text(0.4,24,'Balanced Accuracy: {:.3f}\n'.format(self.BALANCED_ACCURACY))
        plot.text(0.4,22,'Micro Precision: {:.3f}'.format(self.MICRO_PRECISION))
        plot.text(0.4,20,'Micro Recall: {:.3f}'.format(self.MICRO_RECALL))
        plot.text(0.4,18,'Micro F1-score: {:.3f}\n'.format(self.MICRO_F1_SCORE))
        plot.text(0.4,16,'Macro Precision: {:.3f}'.format(self.MACRO_PRECISION))
        plot.text(0.4,14,'Macro Recall: {:.3f}'.format(self.MACRO_RECALL))
        plot.text(0.4,12,'Macro F1-score: {:.3f}\n'.format(self.MACRO_F1_SCORE))
        plot.text(0.4,10,'Weighted Precision: {:.3f}'.format(self.WEIGHTED_PRECISION))
        plot.text(0.4,8,'Weighted Recall: {:.3f}'.format(self.WEIGHTED_RECALL))
        plot.text(0.4,6,'Weighted F1-score: {:.3f}'.format(self.WEIGHTED_F1_SCORE))
        plot.text(0.4,4,'Matthews correlation coefficient: {:.3f}'.format(self.MCC))
        plot.text(0.02,15,'--------------- Classification Report ---------------')
        plot.text(0.02,1,classification_report(self.Y_test, self.Y_predict))
        plot.suptitle(f"Various Result Score for {self.MODEL_NAME}")
        plot.text(0.4,39,'--------------- Imbalanced Report ---------------')
        plot.text(0.4,29,classification_report_imbalanced(self.Y_test, self.Y_predict))
    
    
    # KPI calculation with regression
    def result_report_regression_calculation(self):
        
        # Calculation
        self.DIFF_PREDICT_REAL = self.Y_test - self.Y_predict
        self.MEAN_DIFF_PREDICT_REAL = np.mean(self.DIFF_PREDICT_REAL)
        self.STD_DIFF_PREDICT_REAL = np.std(self.DIFF_PREDICT_REAL)
        self.STANDARD_DIFF_PREDICT_REAL = (self.DIFF_PREDICT_REAL -self. MEAN_DIFF_PREDICT_REAL)/self.STD_DIFF_PREDICT_REAL
        
        # R2 score (coefficient of determination)
        self.R2_SCORE = r2_score(self.Y_test, self.Y_predict)
        
        # Relative Squared Error (RSE)
        self.RSE = np.sum(self.DIFF_PREDICT_REAL*self.DIFF_PREDICT_REAL)/np.sum((self.Y_predict-np.mean(self.Y_predict))*(self.Y_predict-np.mean(self.Y_predict)))
        
        # Mean Squared Error (MSE)
        self.MSE = mean_squared_error(self.Y_test, self.Y_predict)
        
        # Mean Absolute Error (MAE)
        self.MAE = mean_absolute_error(self.Y_test, self.Y_predict)
        
        # Explained variance score
        self.EVC = explained_variance_score(self.Y_test, self.Y_predict)
        
        # Max Error
        self.MAX_ERROR = max_error(self.Y_test, self.Y_predict)
        
        # Root mean squared error (RMSE)
        self.RMSE = root_mean_squared_error(self.Y_test, self.Y_predict)
        
        # # Mean Squared Log Error (MSLE)
        # self.MSLE = mean_squared_log_error(self.Y_test, self.Y_predict)
        
        # # Root Mean Squared Log Error (RMSLE)
        # self.RMSLE = root_mean_squared_log_error(self.Y_test, self.Y_predict)
        
        # Median absolute error 
        self.MEDIAN_ABSOLUTE_ERROR = median_absolute_error(self.Y_test, self.Y_predict)
        
        # # Mean Poisson deviance
        # self.MPD = mean_poisson_deviance(self.Y_test, self.Y_predict)
        
        # # Mean Gamma deviance
        # self.MGD = mean_gamma_deviance(self.Y_test, self.Y_predict)
        
        # Mean Absolute Percentage Error
        self.MAPE = mean_absolute_percentage_error(self.Y_test, self.Y_predict)
        
        # D2 Absolute Error Score
        self.D2 = d2_absolute_error_score(self.Y_test, self.Y_predict)
    
    
    # KPI print with regression
    def result_report_regression_print(self):
        
        print('\n-------------------- Key Metrics --------------------')
        print('\nR2 SCORE: {:.3f} %'.format(100*self.R2_SCORE))
        print('Represent the percentage of observed variation that can be explained by model inputs')
        
        print('\nMaximum Prediction Error: {:.0f}'.format(self.MAX_ERROR))

        print('\nRelative Squared Error (RSE): {:.3f}'.format(self.RSE))
        print('0 means perfectly fit/overfitting')

        print('\nMean Squared Error (MSE): {:.0f} €'.format(self.MSE))
        print('Mean Absolute Error (MAE): {:.0f} €'.format(self.MAE))
        print('Root Mean Squared Error (RMSE): {:.0f} €'.format(self.RMSE))
        print('Mean Absolute Percentage Error (MAPE): {:.3f} %'.format(100*self.MAPE))
        
        # print('\nMean Squared Log Error (MSLE): {:.3f}'.format(self.MSLE))
        # print('Root Mean Squared Log Error (RMSLE): {:.3f}'.format(self.RMSLE))
        
        print('\nExplained variance score: {:.3f}'.format(self.EVC))
        print('D2 Absolute Error Score: {:3f} / 1.0'.format(self.D2))
        print('Median absolute error : {:.0f} €'.format(self.MEDIAN_ABSOLUTE_ERROR))
        # print('Mean Poisson deviance: {:.0f}'.format(self.MPD))
        # print('Mean Gamma deviance: {:.3f}'.format(self.MGD))

    
    # KPI plot with regression
    def result_report_regression_plot(self):
    
        # Plot
        fig,ax = plot.subplots(2)
        ax[0].hist(self.DIFF_PREDICT_REAL, bins = int(self.DIFF_PREDICT_REAL.shape[0]/4))
        ax[0].plot([self.STD_DIFF_PREDICT_REAL,self.STD_DIFF_PREDICT_REAL],[0,int(self.DIFF_PREDICT_REAL.shape[0]/10)], color = "black")
        ax[0].plot([-self.STD_DIFF_PREDICT_REAL,-self.STD_DIFF_PREDICT_REAL],[0,int(self.DIFF_PREDICT_REAL.shape[0]/10)], color = "black")
        ax[0].text(self.STD_DIFF_PREDICT_REAL, int(self.DIFF_PREDICT_REAL.shape[0]/10) + 2, "σ", color = "black")
        ax[0].text(-self.STD_DIFF_PREDICT_REAL, int(self.DIFF_PREDICT_REAL.shape[0]/10) + 2, "σ", color = "black")
        ax[0].plot([2*self.STD_DIFF_PREDICT_REAL,2*self.STD_DIFF_PREDICT_REAL],[0,int(self.DIFF_PREDICT_REAL.shape[0]/10)], color = "orange")
        ax[0].plot([-2*self.STD_DIFF_PREDICT_REAL,-2*self.STD_DIFF_PREDICT_REAL],[0,int(self.DIFF_PREDICT_REAL.shape[0]/10)], color = "orange")
        ax[0].text(2*self.STD_DIFF_PREDICT_REAL, int(self.DIFF_PREDICT_REAL.shape[0]/10) + 2, "2σ", color = "orange")
        ax[0].text(-2*self.STD_DIFF_PREDICT_REAL, int(self.DIFF_PREDICT_REAL.shape[0]/10) + 2, "2σ", color = "orange")
        ax[0].plot([3*self.STD_DIFF_PREDICT_REAL,3*self.STD_DIFF_PREDICT_REAL],[0,int(self.DIFF_PREDICT_REAL.shape[0]/10)], color = "green")
        ax[0].plot([-3*self.STD_DIFF_PREDICT_REAL,-3*self.STD_DIFF_PREDICT_REAL],[0,int(self.DIFF_PREDICT_REAL.shape[0]/10)], color = "green")
        ax[0].text(3*self.STD_DIFF_PREDICT_REAL, int(self.DIFF_PREDICT_REAL.shape[0]/10) + 2, "3σ", color = "green")
        ax[0].text(-3*self.STD_DIFF_PREDICT_REAL, int(self.DIFF_PREDICT_REAL.shape[0]/10) + 2, "3σ", color = "green")
        ax[0].set_title("Histogram on the difference between prediction and reality",)
        ax[1].scatter(self.Y_predict, self.STANDARD_DIFF_PREDICT_REAL)
        ax[1].set_ylim([-10,10])
        ax[1].set_title("Standard deviation regarding predicted values")
        plot.grid()
        
        # Print plot
        plot.figure(figsize = (10,8))
        plot.ylim(1,40)
        plot.text(0.1,38,'-------------------- Key Metrics --------------------')
        plot.text(0.1,36,'R2 SCORE: {:.3f} %'.format(100*self.R2_SCORE))
        plot.text(0.1,35,'Represent the percentage of observed variation that can be explained by model inputs')
        plot.text(0.1,33,'Relative Squared Error (RSE): {:.3f}'.format(self.RSE))
        plot.text(0.1,32,'0 means perfectly fit/overfitting')
        plot.text(0.1,30,'Mean Squared Error (MSE): {:.0f} €'.format(self.MSE))
        plot.text(0.1,29,'Mean Absolute Error (MAE): {:.0f} €'.format(self.MAE))
        plot.text(0.1,28,'Root Mean Squared Error (RMSE): {:.0f} €'.format(self.RMSE))
        plot.text(0.1,27,'Mean Absolute Percentage Error (MAPE): {:.3f} %'.format(100*self.MAPE))
        plot.text(0.1,25,'Maximum prediction error : {:.0f} €'.format(self.MAX_ERROR))
        # plot.text(0.1,23,'Mean Squared Log Error (MSLE): {:.3f}'.format(self.MSLE))
        # plot.text(0.1,22,'Root Mean Squared Log Error (RMSLE): {:.3f}'.format(self.RMSLE))
        plot.text(0.1,20,'Explained variance score: {:.3f} %'.format(100*self.EVC))
        plot.text(0.1,19,'D2 Absolute Error Score: {:3f}'.format(self.D2))
        plot.text(0.1,18,'Median absolute error : {:.0f} €'.format(self.MEDIAN_ABSOLUTE_ERROR))
        # plot.text(0.1,17,'Mean Poisson deviance: {:.0f}'.format(self.MPD))
        # plot.text(0.1,16,'Mean Gamma deviance: {:.3f}'.format(self.MGD))