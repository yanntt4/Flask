/* Creation du canvas */
const canvas = document.querySelector(".monCanevas");
const width = (canvas.width = window.innerWidth);
const height = (canvas.height = window.innerWidth*0.5);
const ctx = canvas.getContext("2d");

/* Fonction pour convertir des degrees en radians */
function degToRad(degrees){
    return (degrees*Math.PI)/180;
}

/* Creation d'un demi cercle comme base du compteur */
ctx.fillStyle = "rgb(34.84356228033175,63.868191415955266,0)";
ctx.beginPath();
ctx.arc(window.innerWidth/2,280,200,degToRad(180),degToRad(0),false);
ctx.fill();

/* Creation d'une ligne pointillee au milieu du demi-cercle */
ctx.fillStyle = "rgb(255,255,255)";
ctx.beginPath();
ctx.setLineDash([5,5]);
ctx.moveTo(window.innerWidth/2, 280);
ctx.lineTo(window.innerWidth/2, 80);
ctx.lineWidth = 2;
ctx.stroke();

/* Creation de l'aguille du compteur */
ctx.strokeStyle = '#4488EE';
ctx.beginPath();
ctx.setLineDash([]);
ctx.moveTo(window.innerWidth/2, 280);
ctx.lineTo(window.innerWidth/2 + 79, 97);
ctx.lineWidth = 2;
ctx.stroke();
ctx.fillStyle = '#4488EE';
ctx.beginPath();
ctx.moveTo(window.innerWidth/2 + 79, 97);
ctx.lineTo(window.innerWidth/2 + 63.43757755387386, 112.04792868503282);
ctx.lineTo(window.innerWidth/2 + 78.63619797265274, 118.6447868593438);
ctx.lineTo(window.innerWidth/2 + 79, 97);
ctx.fill();

/* Creation d'un texte de chaque cote de l'aiguille */
ctx.fillStyle = 'rgb(0,255,0)';
ctx.font = "48px georgia";
ctx.fillText("On time",width/2 + 120,320);
ctx.fillStyle = 'rgb(255,0,0)';
ctx.font = "48px georgia";
ctx.fillText("Delayed",width/2 - 300 ,320);

/* Creation d'un texte au bout de l'aiguille */
ctx.fillStyle = "rgb(34.84356228033175,63.868191415955266,0)";
ctx.font = "48px georgia"
var textString = "63.03 %",
    textWidth = ctx.measureText(textString).width;
ctx.fillText("63.03 %",(width/2) - (textWidth / 2),60)
