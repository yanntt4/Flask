/* Creation du canvas */
var canvas = document.getElementById("canvas1");
const width = (canvas.width = window.innerWidth);
const height = (canvas.height = 350);
const x = 153;
const rect_width = 600;
const triangle_width = 50;
const width_start_rectangle = window.innerWidth/2 - rect_width/2 - triangle_width - 100;
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

/* Creation d'un rectangle comme base */
ctx1.strokeStyle = "rgb(0,0,0)";
ctx1.strokeRect(width_start_rectangle,100,rect_width,100);
ctx1.lineWidth = 2;

/* Creation d'un triangle */
ctx2.strokeStyle = "rgb(0,0,0)";
ctx2.beginPath();
ctx2.moveTo(width_start_rectangle + rect_width, 100);
ctx2.lineTo(window.innerWidth/2 + rect_width/2 + triangle_width, 150);
ctx2.lineTo(width_start_rectangle + rect_width, 200);
ctx2.lineWidth = 1;
ctx2.stroke();
const BI1 = new Image();const BI2 = new Image();const BI3 = new Image();const BI4 = new Image();const BI5 = new Image();const BI6 = new Image();function init() {    BI1.src = "./static/BI1_BACKGROUND.jpg";    BI2.src = "./static/BI2_BACKGROUND.jpg";    BI3.src = "./static/BI3_BACKGROUND.jpg";    BI4.src = "./static/BI4_BACKGROUND.jpg";    BI5.src = "./static/BI5_BACKGROUND.jpg";    BI6.src = "./static/BI6_BACKGROUND.jpg";    window.requestAnimationFrame(draw);}function draw() {    ctx6.drawImage(BI1, width_start_rectangle - 25, 40, 40, 40);    ctx7.drawImage(BI2, width_start_rectangle + 95, 40, 40, 40);    ctx8.drawImage(BI3, width_start_rectangle + 215, 40, 40, 40);    ctx9.drawImage(BI4, width_start_rectangle + 335, 40, 40, 40);    ctx10.drawImage(BI5, width_start_rectangle + 455, 40, 40, 40);    ctx11.drawImage(BI6, width_start_rectangle + 575, 40, 40, 40);    window.requestAnimationFrame(draw);}init()
/* Creation d'une animation pour afficher le prix progressivement */
ctx3.fillStyle = 'rgb(0,0,50)';
ctx3.font = "48px georgia";
var textString = "153379 euros",
    textWidth = ctx3.measureText(textString).width;
var id2 = null;
function myPrice() {
  var price = 0;
   const x2 = 153379;
  id2 = setInterval(frame2, 10);
  function frame2() {
    if (price < x2) { 
      ctx3.clearRect(0,200,3*width/4,100);
      price += getRandomInt(900, 1000);
      var textString2 = `${price} euros`;
          textWidth2 = ctx3.measureText(textString2).width;
      ctx3.fillText(numberWithCommas(textString2),width_start_rectangle + 2*x2/1000 - (textWidth/2),260);
    } else { 
      ctx3.clearRect(0,200,3*width/4,100);
      ctx3.fillText(numberWithCommas(textString),width_start_rectangle + 2*x - (textWidth/2),260);
      clearInterval(id);
    }
  }
}

/* Creation d'une animation pour bouger une barre selon le prix */
var id = null;
function myMove() {
  var pos = 0;
   const x = 153;
  clearInterval(id);
  id = setInterval(frame, 10);
  function frame() {
    if (pos == x) {
      clearInterval(id);
    } else {
      pos++;
      var color = 255 - pos;
      ctx5.strokeStyle = rgb(pos,color/2,color);
      ctx5.beginPath();
      ctx5.moveTo(width_start_rectangle + 2*pos, 100);
      ctx5.lineTo(width_start_rectangle + 2*pos, 200);
      ctx5.lineTo(width_start_rectangle + 2*pos, 100);
      ctx5.lineWidth = 1;
      ctx5.stroke();
    }
  }
}
