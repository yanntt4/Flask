/* Getting element inside select list created in HTML */
var brand_select = document.getElementById("brand");
var model_select = document.getElementById("model");
var fuel_type_select = document.getElementById("fuel_type");
var engine_select = document.getElementById("engine");
var transmission_select = document.getElementById("transmission");

/* Explanations 
On selection, trigger another internet page to get option following selection
Options are sent back with specific format (JSON) to be added to HTML

/* adding model options following brand selection */
brand_select.onchange = function()  {
    brand = brand_select.value;
    fetch('/model/' + brand).then(function(response) {
        response.json().then(function(data) {
            console.log(data.models)
            var optionHTML = "";
            for (var model of data.models) {
                 optionHTML += '<option value="' + model.id + '">' + model.name + '</option>';
             }
            model_select.innerHTML = optionHTML;
        })
    });
}

/* Trigger on model change */
model_select.onchange = function()  {
    model = model_select.value;
    
    /* adding fuel-type options following model selection */
    fetch('/fuel_type/' + model).then(function(response) {
        response.json().then(function(data) {
            console.log(data.models)
            var optionHTML = "";
            for (var fuel_type of data.fuel_types) {
                 optionHTML += '<option value="' + fuel_type.id + '">' + fuel_type.name + '</option>';
             }
             fuel_type_select.innerHTML = optionHTML;
        })
    });

    /* adding engine options following model selection */
    fetch('/engine/' + model).then(function(response) {
        response.json().then(function(data) {
            console.log(data.models)
            var optionHTML = "";
            for (var engine of data.engines) {
                 optionHTML += '<option value="' + engine.id + '">' + engine.name + '</option>';
             }
             engine_select.innerHTML = optionHTML;
        })
    });

    /* adding transmission options following model selection */
    fetch('/transmission/' + model).then(function(response) {
        response.json().then(function(data) {
            console.log(data.models)
            var optionHTML = "";
            for (var transmission of data.transmissions) {
                 optionHTML += '<option value="' + transmission.id + '">' + transmission.name + '</option>';
             }
             transmission_select.innerHTML = optionHTML;
        })
    });
}
