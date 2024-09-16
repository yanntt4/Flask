/* Fonction pour convertir des degres en radians */
function degToRad(degrees){
    return (degrees*Math.PI)/180;
}

/* Fonction pour modifier le style des nombres affiches (style europeen) */
function numberWithCommas(x) {
    return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
 }

/* Fonction pour generer un entier aleatoire entre 2 valeurs */
function getRandomInt(min, max) {
    const minCeiled = Math.ceil(min);
    const maxFloored = Math.floor(max);
    return Math.floor(Math.random() * (maxFloored - minCeiled) + minCeiled);
}

/* Fonction pour changer la couleur avec des valeurs numeriques */
function rgb(r, g, b){
    return "rgb("+r+","+g+","+b+")";
    }