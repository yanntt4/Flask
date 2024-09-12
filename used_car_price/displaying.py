# -*- coding: utf-8 -*-
"""
Created on Fri Nov 17 13:12:32 2023

@author: ythiriet
"""


# Function to create HTML page to ask user data to make prediction
def preparation():


    # Global importation
    from yattag import Doc
    from yattag import indent
    import joblib
    import numpy as np
    
    def html_preparation_creation():

        # Data importation
        ARRAY_DATA_ENCODE_REPLACEMENT = joblib.load("./script/data_replacement/array_data_encode_replacement.joblib")
        NAME_DATA_ENCODE_REPLACEMENT = np.zeros([ARRAY_DATA_ENCODE_REPLACEMENT.shape[0]], dtype = object)
        for i, ARRAY in enumerate(ARRAY_DATA_ENCODE_REPLACEMENT):
            NAME_DATA_ENCODE_REPLACEMENT[i] = ARRAY[0,0]
    
        # Creating HTML
        doc, tag, text, line = Doc().ttl()
    
        # Adding pre-head
        doc.asis('<!DOCTYPE html>')
        doc.asis('<html lang="fr">')
        with tag('head'):
            doc.asis('<meta charset="UTF-8">')
            doc.asis('<meta http-equiv="X-UA-Compatible" content = "IE=edge">')
            doc.asis('<link rel="stylesheet" href="./static/style.css">')
            doc.asis('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">')
            doc.asis('<meta name = "viewport" content="width=device-width, initial-scale = 1.0">')
    
        # Body start
        with tag('body', klass = 'background'):
            with tag('div', klass = "container", id="container"):
                with tag('div', klass = "row"):
                    with tag('div', klass = "col-md-9"):
                        line('h1', "Prix d'une voiture ancienne", klass = "text-center title_s")
                    with tag('div', klass = "col"):
                        doc.asis('<img src="/static/old_car_2.jpg" alt="Car" width=100% height=100% title="Car"/>')
                    
                # Launching prediction
                # with tag('form', method = "POST"):
                with tag('form', action = "{{url_for('treatment')}}", method = "POST", enctype = "multipart/form-data", id = "URL"):
                    text("{{ form.csrf_token }}")

                    # Brand
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"): 
                                with tag('div', klass = "col-md-6"):
                                    line('p', 'Marque du vehicule', klass = "p2")
                                with tag('div', klass = "col-md-4"):
                                    text("{{ form.brand() }}")
    
                    # Model
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"):
                                with tag('div', klass = "col-md-6"):
                                    line('p', 'Modele du vehicule', klass = "p2")
                                with tag('div', klass = "col-md-4"):
                                    text("{{ form.model() }}")
                                    
                    # Vehicle Age
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"):
                                with tag('div', klass = "col-md-6"):
                                    line('p', 'Annee de mise en circulation', klass = "p2")
                                with tag('div', klass = "col", style="text-align:center"):
                                    text("{{ form.model_year() }}")
                    
                    # Milage
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"):
                                with tag('div', klass = "col-md-6"):
                                    line('p', 'Kilometrage du vehicule', klass = "p2")
                                with tag('div', klass = "col", style="text-align:center"):
                                    text("{{ form.milage() }}")
                    
                    # Fuel type
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"): 
                                with tag('div', klass = "col-md-6"):
                                    line('p', 'Carburant utilise', klass = "p2")
                                with tag('div', klass = "col", style="text-align:center"):
                                    text("{{ form.fuel_type() }}")
                    
                    # Engine
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"):
                                with tag('div', klass = "col-md-4"):
                                    line('p', 'Type de moteur', klass = "p2")
                                with tag('select', name = "engine", id="engine"):
                                    text("{{ form.engine() }}")
                    
                    # Transmission
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"):
                                with tag('div', klass = "col-md-4"):
                                    line('p', 'Type de transmission', klass = "p2")
                                with tag('select', name = "transmission", id="transmission"):
                                    text("{{ form.transmission() }}")
                    
                    # Exterior color
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"): 
                                with tag('div', klass = "col-md-6"):
                                    line('p', 'Marque du vehicule', klass = "p2")
                                with tag('div', klass = "col-md-4"):
                                    text("{{ form.ext_col() }}")
                    
                    # Interior color
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"): 
                                with tag('div', klass = "col-md-6"):
                                    line('p', 'Marque du vehicule', klass = "p2")
                                with tag('div', klass = "col-md-4"):
                                    text("{{ form.int_col() }}")
                    
                    # Accident history
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"): 
                                with tag('div', klass = "col-md-6"):
                                    line('p', 'Vehicule accidente ?', klass = "p2")
                                with tag('div', klass = "col", style="text-align:center"):
                                    with tag('div', klass = "row"):
                                        with tag('div', klass = "col-md-4"):
                                            doc.input(name = 'accident', id = "accident", type = 'radio', value = 0, klass = 'radio_text')
                                            text("NON")
                                        with tag('div', klass = "col-md-4"):
                                            doc.input(name = 'accident', id = "accident", type = 'radio', value = 1, klass = 'radio_text')
                                            text("OUI")
                    
                    # Good state if cleaned
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            with tag('div', klass = "row justify-content-around margin_1"): 
                                with tag('div', klass = "col-md-6"):
                                    line('p', 'Etat correct apres nettoyage ?', klass = "p2")
                                with tag('div', klass = "col", style="text-align:center"):
                                    with tag('div', klass = "row"):
                                        with tag('div', klass = "col-md-4"):
                                            doc.input(name = 'clean_title', id = "clean_title", type = 'radio', value = 0, klass = 'radio_text')
                                            text("NON")
                                        with tag('div', klass = "col-md-4"):
                                            doc.input(name = 'clean_title', id = "clean_title", type = 'radio', value = 1, klass = 'radio_text')
                                            text("OUI")
                    
                    # Submit button
                    with tag('div', klass = "row"):
                        with tag('div', klass = "text-center div2"):
                            with tag('button', id = 'submit_button', name = "action", klass="btn btn-primary", value = 'Predict'):
                                text('Predict')
        
            # Script for javascript
            doc.asis('<script src="/static/predict.js"></script>')
                                    
    
        # Saving HTML created
        with open(f"./templates/predict.html", "w") as f:
            f.write(indent(doc.getvalue(), indentation = '    ', newline = '\n', indent_text = True))
            f.close()
    

    # Only used to create the javascript at the beginning. The file created must be modified
    def javascript_preparation_creation():
        
        JS_PREPARATION = ""
        
        JS_PREPARATION += 'var brand_select = document.getElementById("brand");\n'
        JS_PREPARATION += 'var model_select = document.getElementById("model");\n'
        JS_PREPARATION += 'document.write(5 + 6);\n'

        JS_PREPARATION += 'brand_select.onchange = function()  {\n'
        JS_PREPARATION += '    brand = brand_select.value;\n'
            
        JS_PREPARATION += '    fetch("/model/" + brand).then(function(response) {\n'

        JS_PREPARATION += '        response.json().then(function(data) {\n'
        JS_PREPARATION += '            var optionHTML = "";\n'

        JS_PREPARATION += '            for (var model of data.model) {\n'
        JS_PREPARATION += '                optionHTML += "<' + 'option value="' + 'model.id' + '">' + 'model.name' + '"</option>;\n"'
        JS_PREPARATION += '            }\n'

        JS_PREPARATION += '            model_select.innerHTML = optionHTML;\n'
        JS_PREPARATION += '        })\n'
                
        JS_PREPARATION += '    });\n'
        JS_PREPARATION += '}\n'
        

        with open(f"./static/predict.js", "w") as f:
            f.write(JS_PREPARATION)
            f.close()
    

    # Creating HTML
    # javascript_preparation_creation()
    html_preparation_creation()

    


# Function to make prediction and plotting them for the customer
def prediction(CURRENT_DIRECTORY, MODEL_INPUT_HTML, DATA_NAMES_HTML):

    # Global importation
    import joblib
    import numpy as np
    from yattag import Doc
    from yattag import indent
    import math
    import random
    
    # Global init
    RF_MODEL = False
    NN_MODEL = False
    GB_MODEL = False
    XG_MODEL = True
    REGRESSION = True

    # Class creation
    class Data_prediction():
        def __init__(self, MODEL):
            self.ARRAY_DATA_ENCODE_REPLACEMENT = joblib.load("./script/data_replacement/array_data_encode_replacement.joblib")
            self.DATA_NAMES = joblib.load("./script/data_replacement/data_names.joblib")
            
            self.MODEL = MODEL
            
            self.JS_CANVAS = ""

        
        def entry_data_arrangement(self, MODEL_INPUT_HTML, DATA_NAMES_HTML):
            self.MODEL_INPUT = np.zeros([self.DATA_NAMES.shape[0]], dtype = object)
            DATA_NAMES_HTML = np.array(DATA_NAMES_HTML)
            
            for i, DATA_NAME in enumerate(self.DATA_NAMES):
                self.MODEL_INPUT[i] = MODEL_INPUT_HTML[np.where(DATA_NAME == DATA_NAMES_HTML)[0][0]]


        # Turning word into numbers to make predictions
        def entry_data_modification(self):

            for i in range(self.MODEL_INPUT.shape[0]):
                for ARRAY in self.ARRAY_DATA_ENCODE_REPLACEMENT:
                    if self.DATA_NAMES[i] == ARRAY[0,0]:
                        for j in range(ARRAY.shape[0]):
                            if self.MODEL_INPUT[i] == ARRAY[j,-2]:
                                self.MODEL_INPUT[i] = ARRAY[j,-1]


        # Making prediction using model chosen
        def making_prediction(self, REGRESSION):
            print(self.MODEL_INPUT)
            self.PREDICTION = self.MODEL.predict(self.MODEL_INPUT.reshape(1,-1))
            print(self.PREDICTION)
            
            if REGRESSION == False:
                self.PROBA = self.MODEL.predict_proba(self.MODEL_INPUT.reshape(1,-1))
                print(self.PROBA)


        # Creating javascript using prediction
        def javascript_result_creation(self):
            
            # Creating Canvas to plot graphic
            self.JS_CANVAS += '/* Creation du canvas */\n'
            self.JS_CANVAS += 'var canvas = document.getElementById("canvas1");\n'
            self.JS_CANVAS += 'const width = (canvas.width = window.innerWidth);\n'
            self.JS_CANVAS += 'const height = (canvas.height = 500);\n'
            self.JS_CANVAS += f'const x = {int(self.PREDICTION[0]/1000)};\n'
            self.JS_CANVAS += "canvas.style.position = 'relative';\n"
            self.JS_CANVAS += "canvas.style.zIndex = 1;"
            for i in range(1,7):
                self.JS_CANVAS += f'var ctx{i} = canvas.getContext("2d");\n'
            
            # Function to display correct format for number
            self.JS_CANVAS += '\n/* Fonction pour modifier le style des nombres affichés */\n'
            self.JS_CANVAS += 'function numberWithCommas(x) {\n'
            self.JS_CANVAS += '   return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");\n'
            self.JS_CANVAS += '}\n'
            
            # Function to display correct format for number
            self.JS_CANVAS += '\n/* Fonction pour generer un entier aleatoire entre 2 valeurs */\n'
            self.JS_CANVAS += 'function getRandomInt(min, max) {\n'
            self.JS_CANVAS += '    const minCeiled = Math.ceil(min);\n'
            self.JS_CANVAS += '    const maxFloored = Math.floor(max);\n'
            self.JS_CANVAS += '    return Math.floor(Math.random() * (maxFloored - minCeiled) + minCeiled);\n'
            self.JS_CANVAS += '}\n'
            
            # Function to change color
            self.JS_CANVAS += "\n/* Creation d'une fonction pour changer la couleur */\n"
            self.JS_CANVAS += 'function rgb(r, g, b){\n'
            self.JS_CANVAS += 'return "rgb("+r+","+g+","+b+")";\n'
            self.JS_CANVAS += '}\n'
            
            # Function to draw image with rotation
            self.JS_CANVAS += "\n/* Creation d'une fonction pour faire tourner une seule image dans un canvas */\n"
            self.JS_CANVAS += 'function drawImageRot(ctx,img,x,y,width,height,deg){\n'
            self.JS_CANVAS += '    /* Store the current context state (i.e. rotation, translation etc..) */\n'
            self.JS_CANVAS += '    ctx.save();\n'
            self.JS_CANVAS += '    /* Convert degrees to radian  */\n'
            self.JS_CANVAS += '    var rad = deg * Math.PI / 180;\n'
            self.JS_CANVAS += '    /* Set the origin to the center of the image */\n'
            self.JS_CANVAS += '    ctx.translate(x + width / 2, y + height / 2);\n'
            self.JS_CANVAS += '    /* Rotate the canvas around the origin */\n'
            self.JS_CANVAS += '    ctx.rotate(rad);\n'
            self.JS_CANVAS += '    /* Draw the image */\n'
            self.JS_CANVAS += '    ctx.drawImage(img,width / 2 * (-1),height / 2 * (-1),width,height);\n'
            self.JS_CANVAS += '    /* Restore canvas state as saved from above */\n'
            self.JS_CANVAS += '    ctx.restore();\n'
            self.JS_CANVAS += '}\n'
            
            # Road construction functions
            self.JS_CANVAS += "\n/* Fonction pour creer une route en vue du dessus */\n"
            self.JS_CANVAS += 'function road_creation(ctx) {\n'
            
            self.JS_CANVAS += '    ctx.fillStyle = "rgb(50,50,50)";\n'
            self.JS_CANVAS += '    ctx.fillRect(0, 50, 3*width/4, 150);\n'
            self.JS_CANVAS += '    ctx.save();\n'
            self.JS_CANVAS += '    ctx.fillStyle = "rgb(255,255,255)";\n'
            self.JS_CANVAS += '    ctx.fillRect(0, 200, 3*width/4, 5);\n'
            self.JS_CANVAS += '    ctx.save();\n'
            self.JS_CANVAS += '    ctx.fillStyle = "rgb(20,20,20)";\n'
            self.JS_CANVAS += '    ctx.fillRect(0, 205, 3*width/4, 100);\n'
            self.JS_CANVAS += '    ctx.save();\n'
            self.JS_CANVAS += '    ctx.fillStyle = "rgb(255,255,255)";\n'
            self.JS_CANVAS += '    ctx.fillRect(0, 305, 3*width/4, 5);\n'
            self.JS_CANVAS += '    ctx.save();\n'
            self.JS_CANVAS += '    ctx.fillStyle = "rgb(50,50,50)";\n'
            self.JS_CANVAS += '    ctx.fillRect(0, 310, 3*width/4, 150);\n'
            self.JS_CANVAS += '    ctx.save();\n'
            for i in range(19):
                self.JS_CANVAS += '    ctx.fillStyle = "rgb(255,255,51)";\n'
                self.JS_CANVAS += f'    ctx.fillRect({10 + 60*i}, 250, 50, 10);\n'
                self.JS_CANVAS += '    ctx.save();\n'
            self.JS_CANVAS += '    ctx.restore();\n'
            self.JS_CANVAS += '}\n'
            
            # Function to animate the car
            self.JS_CANVAS += "\n/* Fonction pour creer une animation de vehicule */\n"
            self.JS_CANVAS += 'function car_animation() {\n'
            self.JS_CANVAS += '    const x_pos = 0;\n'
            self.JS_CANVAS += '    const y_pos = 200;\n'
            self.JS_CANVAS += f'    const predict_price = {int(self.PREDICTION)};\n'
            self.JS_CANVAS += '    const car = new Image(30,25);\n'
            self.JS_CANVAS += '    const wheel_forward = new Image(30,25);\n'
            self.JS_CANVAS += '    const wheel_backward = new Image(30,25);\n'
            self.JS_CANVAS += '    id2 = setInterval(frame2,30);\n'
            self.JS_CANVAS += '    var metres = 0;\n'
            self.JS_CANVAS += '    function frame2() {\n'
            self.JS_CANVAS += '        if (metres < 900) {\n'
            self.JS_CANVAS += '            metres += 8;\n'
            self.JS_CANVAS += '            road_creation(ctx2);\n'
            self.JS_CANVAS += '            car.onload = function() {\n'
            self.JS_CANVAS += '                ctx3.drawImage(car,x_pos + metres,y_pos,200,100);'
            self.JS_CANVAS += '            };\n'
            self.JS_CANVAS += '            wheel_forward.onload = function() {\n'
            self.JS_CANVAS += '                drawImageRot(ctx4,wheel_forward,x_pos + 18 + metres,y_pos + 50,40,40,5*metres);\n'
            self.JS_CANVAS += '            };\n'
            self.JS_CANVAS += '            wheel_backward.onload = function() {\n'
            self.JS_CANVAS += '                drawImageRot(ctx5,wheel_backward,x_pos + 140 + metres,y_pos + 50,40,40,5*metres);\n'
            self.JS_CANVAS += '            };\n'
            self.JS_CANVAS += '            car.src = "/static/car_animation.png";\n'
            self.JS_CANVAS += '            wheel_forward.src = "/static/wheel_animation.png";\n'
            self.JS_CANVAS += '            wheel_backward.src = "/static/wheel_animation.png";\n'
            self.JS_CANVAS += '        } else {\n'
            self.JS_CANVAS += '            ctx6.font = "bold 25px verdana, sans-serif";\n'
            self.JS_CANVAS += '            ctx6.fillStyle = rgb(255,255,255);\n'
            self.JS_CANVAS += '            ctx6.fillText(`${numberWithCommas(predict_price)} euros`,x_pos + 10 + metres, y_pos + 180);\n'
            self.JS_CANVAS += '        }\n'
            self.JS_CANVAS += '    }\n'
            self.JS_CANVAS += '}\n'
    
            # Writing Javascript into a file
            with open("./static/canevas.js","w") as f:
                f.write(self.JS_CANVAS)
    

        # Creating html
        def html_result_creation(self, CURRENT_DIRECTORY):

            # Creating HTML
            doc, tag, text, line = Doc(defaults = {'Month': 'Fevrier'}).ttl()

            doc.asis('<!DOCTYPE html>')
            doc.asis('<html lang="fr">')
            with tag('head'):
                doc.asis('<meta charset="UTF-8">')
                doc.asis('<meta http-equiv="X-UA-Compatible" content = "IE=edge">')
                doc.asis('<link rel="stylesheet" href="./static/style.css">')
                doc.asis('<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH" crossorigin="anonymous">')
                doc.asis('<meta name = "viewport" content="width=device-width, initial-scale = 1.0">')

            # Body start
            with tag('body', klass = 'background'):
                with tag('div', klass = "container"):
                    with tag('div', klass = "row"):
                        with tag('div', klass = "col-md-9"):
                            line('h1', "Prix d'une voiture ancienne", klass = "text-center title_s")
                        with tag('div', klass = "col"):
                            doc.asis('<img src="/static/old_car_2.jpg" alt="Car" width=100% height=100% title="Car"/>')

                    with tag('div', klass="col"):
                        with tag('canvas', id = "canvas1", width="540", height="600"):
                            text("")
                    
                            # Script for canvas
                            doc.asis('<script src="/static/canevas.js"></script>')
                        
                        # Launching script when arriving on the page
                        with tag('script', type="text/javascript"):
                            text('car_animation();')
                
                # Button to go back to previous page
                with tag('form', action = "{{url_for('predict')}}", method = "GET", enctype = "multipart/form-data"):
                    with tag('div', klass = "text-center"):
                        with tag('button', id = 'submit_button', name = "action", klass="btn btn-primary", value = 'Go back to previous page'):
                            line('p1', '')
                            text('Go back to previous page')

            # Saving HTML
            with open("./templates/result.html", "w") as f:
                f.write(indent(doc.getvalue(), indentation = '    ', newline = '\n', indent_text = True))
                f.close()

    # Loading models
    if RF_MODEL == True:
        with open("./script/models/rf_model.sav", 'rb') as f:
            MODEL = joblib.load(f)
    elif NN_MODEL == True:
        with open("./script/models/nn_model.sav", 'rb') as f:
            MODEL = joblib.load(f)
    elif GB_MODEL == True:
        with open("./script/models/gb_model.sav", 'rb') as f:
            MODEL = joblib.load(f)
    elif XG_MODEL == True:
        with open("./script/models/xg_model.sav", 'rb') as f:
            MODEL = joblib.load(f)

    # Personnalized prediction
    global_data_prediction = Data_prediction(MODEL)
    global_data_prediction.entry_data_arrangement(MODEL_INPUT_HTML, DATA_NAMES_HTML)
    global_data_prediction.entry_data_modification()

    global_data_prediction.making_prediction(REGRESSION)
    global_data_prediction.javascript_result_creation()
    global_data_prediction.html_result_creation(CURRENT_DIRECTORY)
