# -*- coding: utf-8 -*-
"""
Created on Thu Nov 23 10:00:27 2023

@author: ythiriet
"""

# Global librairies
from flask import Flask, request, render_template, send_file, redirect, url_for, jsonify
from flask_httpauth import HTTPBasicAuth
from werkzeug.security import generate_password_hash, check_password_hash
import random
import joblib
import os
import sys
import numpy as np
import pandas as pd
import shutil
import openpyxl
import os
from zipfile import ZipFile

from flask_wtf import FlaskForm 
from wtforms import StringField, PasswordField, BooleanField , FieldList
from wtforms import DecimalField, RadioField, SelectField, TextAreaField, FileField 
from wtforms.validators import InputRequired 

# Loading list possibilities
ARRAY_DATA_ENCODE_REPLACEMENT = joblib.load("./script/data_replacement/array_data_encode_replacement.joblib")
NAME_DATA_ENCODE_REPLACEMENT = np.zeros([ARRAY_DATA_ENCODE_REPLACEMENT.shape[0]], dtype = object)
for i, ARRAY in enumerate(ARRAY_DATA_ENCODE_REPLACEMENT):
    NAME_DATA_ENCODE_REPLACEMENT[i] = ARRAY[0,0]

BRANDS = ARRAY_DATA_ENCODE_REPLACEMENT[np.where(NAME_DATA_ENCODE_REPLACEMENT == "brand")[0][0]][:,-2:]
MODELS = ARRAY_DATA_ENCODE_REPLACEMENT[np.where(NAME_DATA_ENCODE_REPLACEMENT == "model")[0][0]][:,-2]
FUEL_TYPES = ARRAY_DATA_ENCODE_REPLACEMENT[np.where(NAME_DATA_ENCODE_REPLACEMENT == "fuel_type")[0][0]][:,-2]
ENGINES = ARRAY_DATA_ENCODE_REPLACEMENT[np.where(NAME_DATA_ENCODE_REPLACEMENT == "engine")[0][0]][:,-2]
TRANSMISSIONS = ARRAY_DATA_ENCODE_REPLACEMENT[np.where(NAME_DATA_ENCODE_REPLACEMENT == "transmission")[0][0]][:,-2]
EXTER_COLORS = ARRAY_DATA_ENCODE_REPLACEMENT[np.where(NAME_DATA_ENCODE_REPLACEMENT == "ext_col")[0][0]][:,-2:]
INTER_COLORS = ARRAY_DATA_ENCODE_REPLACEMENT[np.where(NAME_DATA_ENCODE_REPLACEMENT == "int_col")[0][0]][:,-2:]

BRANDS = np.flip(BRANDS,1).tolist()
EXTER_COLORS = np.flip(EXTER_COLORS,1).tolist()
INTER_COLORS = np.flip(INTER_COLORS,1).tolist()

UNIQUE_BRAND_MODEL = joblib.load("./script/data_linked/unique_brand_model.joblib")
UNIQUE_FUEL_TYPE_MODEL = joblib.load("./script/data_linked/unique_model_fuel_type.joblib")
UNIQUE_ENGINE_MODEL = joblib.load("./script/data_linked/unique_model_engine.joblib")
UNIQUE_TRANSMISSION_MODEL = joblib.load("./script/data_linked/unique_model_transmission.joblib")


# WTF class definition
class Usedcarpriceform(FlaskForm):
    brand = SelectField('brand', choices=BRANDS)
    model = SelectField('model', choices=[])
    model_year = DecimalField('model_year', validators=[InputRequired()])
    milage = DecimalField('milage', validators=[InputRequired()])
    fuel_type = SelectField('fuel-type', choices=[])
    engine = SelectField('engine', choices=[])
    transmission = SelectField('transmission', choices=[])
    ext_col = SelectField('ext_col', choices=EXTER_COLORS)
    int_col = SelectField('int_col', choices=INTER_COLORS)
    accident = RadioField('accident', choices = [('NON','0'),('OUI','1')])
    clean_title = RadioField('clean_title', choices = [('NON','0'),('OUI','1')])


# Local
CURRENT_DIRECTORY = os.path.dirname(os.path.abspath(__file__))
sys.path.append(f"{CURRENT_DIRECTORY}/script/")
import displaying

# App object creation
app = Flask(__name__, static_url_path='/static')
app.config['SECRET_KEY'] = 'secretkey'

# Security object creation
auth = HTTPBasicAuth()

# Password authorized
user = "admin"
pw = "admin"
users = {
    user: generate_password_hash(pw)
}

# Decorator with the verify password function
@auth.verify_password
def verify_password(username, password):
    if username in users:
        return check_password_hash(users.get(username), password)
    return False


# Drop Off page
@app.route("/used_car_price", methods = ["GET", "POST"])
@auth.login_required
def predict():

    # Class definition
    usedcarpriceform = Usedcarpriceform()

    # Creating HTML and JAVASCRIPT
    displaying.preparation()

    # Rendering HTML page
    return render_template("predict.html", form=usedcarpriceform)


# Page to create new options for model following branch
# Page triggered by javascript function (fetch)
@app.route('/model/<brand>')
def brand_select(brand):

    brand = BRANDS[int(brand)][1]

    models = pd.DataFrame(UNIQUE_BRAND_MODEL[np.where(UNIQUE_BRAND_MODEL[:,0] == brand)[0][0],[1]])
    models["index"] = models.index
    models = np.array(models)[0][0]
    models = np.sort(models, axis = 0)
    models = models.tolist()
    global modelArray
    modelArray = []

    for i, model in enumerate(models):
        modelObj = {}
        modelObj['id'] = i
        modelObj['name'] = model
        modelArray.append(modelObj)

    return jsonify({'models' : modelArray})


# Page to create new options for fuel-type following model
# Page triggered by Javascript function (fetch)
@app.route('/fuel_type/<model>')
def model_fuel_type_select(model):

    model = modelArray[int(model)]['name']

    fuel_types = pd.DataFrame(UNIQUE_FUEL_TYPE_MODEL[np.where(UNIQUE_FUEL_TYPE_MODEL[:,0] == model)[0][0],[1]])
    fuel_types["index"] = fuel_types.index
    fuel_types = np.array(fuel_types)[0][0]
    fuel_types = np.sort(fuel_types, axis = 0)
    fuel_types = fuel_types.tolist()
    fuel_typeArray = []

    for i, fuel_type in enumerate(fuel_types):
        fuel_typeObj = {}
        fuel_typeObj['id'] = i
        fuel_typeObj['name'] = fuel_type
        fuel_typeArray.append(fuel_typeObj)

    return jsonify({'fuel_types' : fuel_typeArray})


# Page to create new options for engine following model
# Page triggered by Javascript function (fetch)
@app.route('/engine/<model>')
def model_engine_select(model):

    model = modelArray[int(model)]['name']

    engines = pd.DataFrame(UNIQUE_ENGINE_MODEL[np.where(UNIQUE_ENGINE_MODEL[:,0] == model)[0][0],[1]])
    engines["index"] = engines.index
    engines = np.array(engines)[0][0]
    engines = np.sort(engines, axis = 0)
    engines = engines.tolist()
    engineArray = []

    for i, engine in enumerate(engines):
        engineObj = {}
        engineObj['id'] = i
        engineObj['name'] = engine
        engineArray.append(engineObj)

    return jsonify({'engines' : engineArray})


# Page to create new options for transmission following model
# Page triggered by Javascript function (fetch)
@app.route('/transmission/<model>')
def model_transmission_select(model):

    model = modelArray[int(model)]['name']

    transmissions = pd.DataFrame(UNIQUE_TRANSMISSION_MODEL[np.where(UNIQUE_TRANSMISSION_MODEL[:,0] == model)[0][0],[1]])
    transmissions["index"] = transmissions.index
    transmissions = np.array(transmissions)[0][0]
    transmissions = np.sort(transmissions, axis = 0)
    transmissions = transmissions.tolist()
    transmissionArray = []

    for i, transmission in enumerate(transmissions):
        transmissionObj = {}
        transmissionObj['id'] = i
        transmissionObj['name'] = transmission
        transmissionArray.append(transmissionObj)

    return jsonify({'transmissions' : transmissionArray})


# Page Getting Files Uploaded
@app.route("/treatment", methods = ["GET", "POST"])
def treatment():

    # Init
    DATA_NAMES = ["brand", "model", "model_year", "milage",
                  "fuel_type","engine","transmission",
                  "ext_col","int_col","accident","clean_title"]
    MODEL_INPUT = np.zeros([len(DATA_NAMES)], dtype = object)

    if request.method == "POST":

        for i, NAME in enumerate(DATA_NAMES):
            # Getting infos from predict page
            try:
                MODEL_INPUT[i] = request.form[NAME]
            except Exception as e:
                print(e)
                return render_template("Erreur.html")

        # Making prediction
        displaying.prediction(CURRENT_DIRECTORY, MODEL_INPUT, DATA_NAMES)
        return render_template("result.html")

    # Issue when collecting data
    return render_template("Erreur.html")

# Launching the Server
if __name__ == "__main__":
    app.run(debug=True)
