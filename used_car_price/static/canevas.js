/* Creation du canvas */
var canvas = document.getElementById("canvas1");
const width = (canvas.width = window.innerWidth);
const height = (canvas.height = 500);
const x = 20;
canvas.style.position = 'relative';
canvas.style.zIndex = 1;var ctx1 = canvas.getContext("2d");
var ctx2 = canvas.getContext("2d");
var ctx3 = canvas.getContext("2d");
var ctx4 = canvas.getContext("2d");
var ctx5 = canvas.getContext("2d");
var ctx6 = canvas.getContext("2d");

/* Fonction pour modifier le style des nombres affichés */
function numberWithCommas(x) {
   return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}

/* Fonction pour generer un entier aleatoire entre 2 valeurs */
function getRandomInt(min, max) {
    const minCeiled = Math.ceil(min);
    const maxFloored = Math.floor(max);
    return Math.floor(Math.random() * (maxFloored - minCeiled) + minCeiled);
}

/* Creation d'une fonction pour changer la couleur */
function rgb(r, g, b){
return "rgb("+r+","+g+","+b+")";
}

/* Creation d'une fonction pour faire tourner une seule image dans un canvas */
function drawImageRot(ctx,img,x,y,width,height,deg){
    /* Store the current context state (i.e. rotation, translation etc..) */
    ctx.save();
    /* Convert degrees to radian  */
    var rad = deg * Math.PI / 180;
    /* Set the origin to the center of the image */
    ctx.translate(x + width / 2, y + height / 2);
    /* Rotate the canvas around the origin */
    ctx.rotate(rad);
    /* Draw the image */
    ctx.drawImage(img,width / 2 * (-1),height / 2 * (-1),width,height);
    /* Restore canvas state as saved from above */
    ctx.restore();
}

/* Fonction pour creer une route en vue du dessus */
function road_creation(ctx) {
    ctx.fillStyle = "rgb(50,50,50)";
    ctx.fillRect(0, 50, 3*width/4, 150);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,255)";
    ctx.fillRect(0, 200, 3*width/4, 5);
    ctx.save();
    ctx.fillStyle = "rgb(20,20,20)";
    ctx.fillRect(0, 205, 3*width/4, 100);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,255)";
    ctx.fillRect(0, 305, 3*width/4, 5);
    ctx.save();
    ctx.fillStyle = "rgb(50,50,50)";
    ctx.fillRect(0, 310, 3*width/4, 150);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(10, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(70, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(130, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(190, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(250, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(310, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(370, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(430, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(490, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(550, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(610, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(670, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(730, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(790, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(850, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(910, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(970, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(1030, 250, 50, 10);
    ctx.save();
    ctx.fillStyle = "rgb(255,255,51)";
    ctx.fillRect(1090, 250, 50, 10);
    ctx.save();
    ctx.restore();
}

/* Fonction pour creer une animation de vehicule */
function car_animation() {
    const x_pos = 0;
    const y_pos = 200;
    const predict_price = 20052;
    const car = new Image(30,25);
    const wheel_forward = new Image(30,25);
    const wheel_backward = new Image(30,25);
    id2 = setInterval(frame2,30);
    var metres = 0;
    function frame2() {
        if (metres < 900) {
            metres += 8;
            road_creation(ctx2);
            car.onload = function() {
                ctx3.drawImage(car,x_pos + metres,y_pos,200,100);            };
            wheel_forward.onload = function() {
                drawImageRot(ctx4,wheel_forward,x_pos + 18 + metres,y_pos + 50,40,40,5*metres);
            };
            wheel_backward.onload = function() {
                drawImageRot(ctx5,wheel_backward,x_pos + 140 + metres,y_pos + 50,40,40,5*metres);
            };
            car.src = "/static/car_animation.png";
            wheel_forward.src = "/static/wheel_animation.png";
            wheel_backward.src = "/static/wheel_animation.png";
        } else {
            ctx6.font = "bold 25px verdana, sans-serif";
            ctx6.fillStyle = rgb(255,255,255);
            ctx6.fillText(`${numberWithCommas(predict_price)} euros`,x_pos + 10 + metres, y_pos + 180);
        }
    }
}
