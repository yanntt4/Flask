/* Creation du canvas */
var canvas = document.getElementById("canvas1");
const width = (canvas.width = window.innerWidth);
const height = (canvas.height = 500);
const x = 0;
canvas.style.position = 'relative';
canvas.style.zIndex = 1;var ctx0 = canvas.getContext("2d");
var ctx1 = canvas.getContext("2d");
var ctx2 = canvas.getContext("2d");
var ctx3 = canvas.getContext("2d");
var ctx4 = canvas.getContext("2d");
var ctx5 = canvas.getContext("2d");
var ctx6 = canvas.getContext("2d");
var ctx7 = canvas.getContext("2d");
var ctx8 = canvas.getContext("2d");
var ctx9 = canvas.getContext("2d");
var ctx10 = canvas.getContext("2d");
var ctx11 = canvas.getContext("2d");

/* Fonction pour modifier le style des nombres affich�s */
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

/* Fonction pour creer une formule 1 orientee vers la gauche */
function create_auto_drawing_left(wheel_position_x, wheel_position_y) {
    ctx2.strokeStyle = "rgb(0,0,0)";
    ctx2.beginPath();
    ctx2.arc(wheel_position_x,wheel_position_y,12,0,2*Math.PI,false);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.arc(wheel_position_x + 120,wheel_position_y,12,0,2*Math.PI,false);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.moveTo(wheel_position_x - 12,wheel_position_y + 4);
    ctx2.lineTo(wheel_position_x - 50,wheel_position_y + 4);
    ctx2.bezierCurveTo(wheel_position_x - 25,wheel_position_y - 16,wheel_position_x,wheel_position_y - 19,wheel_position_x + 25,wheel_position_y - 21);
    ctx2.bezierCurveTo(wheel_position_x + 55,wheel_position_y - 20,wheel_position_x + 65,wheel_position_y - 18,wheel_position_x + 75,wheel_position_y - 16);
    ctx2.bezierCurveTo(wheel_position_x + 85,wheel_position_y - 14,wheel_position_x + 95,wheel_position_y - 12,wheel_position_x + 105,wheel_position_y - 11);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.moveTo(wheel_position_x + 13,wheel_position_y + 4);
    ctx2.lineTo(wheel_position_x + 102,wheel_position_y + 4);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.moveTo(wheel_position_x + 55,wheel_position_y - 20);
    ctx2.lineTo(wheel_position_x + 55,wheel_position_y - 35);
    ctx2.bezierCurveTo(wheel_position_x + 59,wheel_position_y - 37,wheel_position_x + 64,wheel_position_y - 38.3,wheel_position_x + 67,wheel_position_y - 38);
    ctx2.bezierCurveTo(wheel_position_x + 70,wheel_position_y - 38,wheel_position_x + 80,wheel_position_y - 36,wheel_position_x + 90,wheel_position_y - 34);
    ctx2.bezierCurveTo(wheel_position_x + 100,wheel_position_y - 31,wheel_position_x + 120,wheel_position_y - 29,wheel_position_x + 140,wheel_position_y - 28);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.moveTo(wheel_position_x + 130,wheel_position_y - 6);
    ctx2.bezierCurveTo(wheel_position_x + 145,wheel_position_y - 11,wheel_position_x + 147,wheel_position_y - 22,wheel_position_x + 150,wheel_position_y - 37);
    ctx2.lineTo(wheel_position_x + 130,wheel_position_y - 37);
    ctx2.stroke();
    ctx2.lineWidth = 1;
}

/* Fonction pour creer une formule 1 orientee vers la droite */
function create_auto_drawing_right(wheel_position_x, wheel_position_y) {
    ctx2.strokeStyle = "rgb(0,0,0)";
    ctx2.beginPath();
    ctx2.arc(wheel_position_x,wheel_position_y,12,0,2*Math.PI,false);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.arc(wheel_position_x - 120,wheel_position_y,12,0,2*Math.PI,false);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.moveTo(wheel_position_x + 12,wheel_position_y + 4);
    ctx2.lineTo(wheel_position_x + 50,wheel_position_y + 4);
    ctx2.bezierCurveTo(wheel_position_x + 25,wheel_position_y - 16,wheel_position_x,wheel_position_y - 19,wheel_position_x - 25,wheel_position_y - 21);
    ctx2.bezierCurveTo(wheel_position_x - 55,wheel_position_y - 20,wheel_position_x - 65,wheel_position_y - 18,wheel_position_x - 75,wheel_position_y - 16);
    ctx2.bezierCurveTo(wheel_position_x - 85,wheel_position_y - 14,wheel_position_x - 95,wheel_position_y - 12,wheel_position_x - 105,wheel_position_y - 11);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.moveTo(wheel_position_x - 13,wheel_position_y + 4);
    ctx2.lineTo(wheel_position_x - 102,wheel_position_y + 4);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.moveTo(wheel_position_x - 55,wheel_position_y - 20);
    ctx2.lineTo(wheel_position_x - 55,wheel_position_y - 35);
    ctx2.bezierCurveTo(wheel_position_x - 59,wheel_position_y - 37,wheel_position_x - 64,wheel_position_y - 38.3,wheel_position_x - 67,wheel_position_y - 38);
    ctx2.bezierCurveTo(wheel_position_x - 70,wheel_position_y - 38,wheel_position_x - 80,wheel_position_y - 36,wheel_position_x - 90,wheel_position_y - 34);
    ctx2.bezierCurveTo(wheel_position_x - 100,wheel_position_y - 31,wheel_position_x - 120,wheel_position_y - 29,wheel_position_x - 140,wheel_position_y - 28);
    ctx2.stroke();
    ctx2.beginPath();
    ctx2.moveTo(wheel_position_x - 130,wheel_position_y - 6);
    ctx2.bezierCurveTo(wheel_position_x - 145,wheel_position_y - 11,wheel_position_x - 147,wheel_position_y - 22,wheel_position_x - 150,wheel_position_y - 37);
    ctx2.lineTo(wheel_position_x - 130,wheel_position_y - 37);
    ctx2.stroke();
    ctx2.lineWidth = 1;
}

/* Creation d'une animation pour faire bouger la voiture */
var id = null;
function myMove() {
  var pos = 150;
  var proba = 0.5032687783241272;
  var proba_percent = 50;
   const x = pos + 1000*proba;
  clearInterval(id);
  id = setInterval(frame, 10);
  function frame() {
    if (pos > x) {
      ctx2.clearRect(0,0,width,190);
      create_auto_drawing_right(pos,150);
      ctx2.font = "38px georgia";
     ctx2.strokeText(`${proba_percent} %`,pos - 100,90);
    } else {
      ctx2.clearRect(0,0,width,190);
      pos += getRandomInt(1,5);
      create_auto_drawing_right(pos,150);
    }
  }
}
ctx3.strokeStyle = "rgb(0,0,0)";
var line_height = 200;
ctx3.lineWidth = 1;
ctx4.lineWidth = 3;
ctx5.font = "28px georgia";
ctx3.beginPath();
ctx3.moveTo(100,line_height);
ctx3.lineTo(1200,line_height);
ctx3.lineTo(1190,line_height - 10);
ctx3.lineTo(1200,line_height);
ctx3.lineTo(1190,line_height + 10);
ctx5.strokeText("[%]",1200,line_height + 33);
ctx5.strokeText("Pourcentage de chance que le client soit interesse",width/6,line_height + 100);
ctx3.stroke();
ctx4.beginPath();
ctx4.moveTo(200,line_height + 10);
ctx4.lineTo(200,line_height - 10);
ctx5.strokeText(10.0,185,line_height + 33);
ctx4.stroke();
ctx4.beginPath();
ctx4.moveTo(300,line_height + 10);
ctx4.lineTo(300,line_height - 10);
ctx5.strokeText(20.0,285,line_height + 33);
ctx4.stroke();
ctx4.beginPath();
ctx4.moveTo(400,line_height + 10);
ctx4.lineTo(400,line_height - 10);
ctx5.strokeText(30.0,385,line_height + 33);
ctx4.stroke();
ctx4.beginPath();
ctx4.moveTo(500,line_height + 10);
ctx4.lineTo(500,line_height - 10);
ctx5.strokeText(40.0,485,line_height + 33);
ctx4.stroke();
ctx4.beginPath();
ctx4.moveTo(600,line_height + 10);
ctx4.lineTo(600,line_height - 10);
ctx5.strokeText(50.0,585,line_height + 33);
ctx4.stroke();
ctx4.beginPath();
ctx4.moveTo(700,line_height + 10);
ctx4.lineTo(700,line_height - 10);
ctx5.strokeText(60.0,685,line_height + 33);
ctx4.stroke();
ctx4.beginPath();
ctx4.moveTo(800,line_height + 10);
ctx4.lineTo(800,line_height - 10);
ctx5.strokeText(70.0,785,line_height + 33);
ctx4.stroke();
ctx4.beginPath();
ctx4.moveTo(900,line_height + 10);
ctx4.lineTo(900,line_height - 10);
ctx5.strokeText(80.0,885,line_height + 33);
ctx4.stroke();
ctx4.beginPath();
ctx4.moveTo(1000,line_height + 10);
ctx4.lineTo(1000,line_height - 10);
ctx5.strokeText(90.0,985,line_height + 33);
ctx4.stroke();
ctx4.beginPath();
ctx4.moveTo(1100,line_height + 10);
ctx4.lineTo(1100,line_height - 10);
ctx5.strokeText(100.0,1085,line_height + 33);
ctx4.stroke();
