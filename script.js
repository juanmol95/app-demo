function empezar() {

    document.getElementById("pantalla-inicio").style.display = "none";

    document.getElementById("pantalla-opciones").style.display = "flex";

}


function irARegistrar() {

    document.getElementById("pantalla-opciones").style.display = "none";

    document.getElementById("pantalla-registrar").style.display = "flex";

}


function irABuscar() {

    document.getElementById("pantalla-opciones").style.display = "none";

    document.getElementById("pantalla-buscar").style.display = "flex";

}


function volverAOpciones() {

    document.getElementById("pantalla-registrar").style.display = "none";

    document.getElementById("pantalla-buscar").style.display = "none";

    document.getElementById("pantalla-opciones").style.display = "flex";

}
