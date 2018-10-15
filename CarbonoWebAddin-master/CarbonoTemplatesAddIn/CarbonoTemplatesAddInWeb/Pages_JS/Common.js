(function () {
    "use strict";


    // A função inicializar deverá ser executada cada vez que uma nova página for carregada.
    Office.initialize = function (reason) {
        $(document).ready(function () {

        });
    };



})();
function abrirPopup() {
    Office.context.ui.displayDialogAsync('https://microsoft-carbonotemplatesaddin.carbonocorporate.com/Pages/Info.html', { height: 20, width: 47 }); //https://microsoft-carbonotemplatesaddin.carbonocorporate.com/Pages/About.html
}

function TemplateDownload(name) {
    var x = new XMLHttpRequest();
    x.open("GET", "../Templates/" + name, true);
    x.responseType = 'blob';
    x.onload = function (e) { download(x.response, name, "Excel Template for Carbono"); };
    x.send();
}

//Basic
$(document).ready(function () {
    $('.animsition').animsition();
});

 //Data options
//$(document).ready(function () {
//    var $animsition = $('.animsition');
//    $animsition
//        .animsition()
//        .one('animsition.inStart', function () {
            
//            console.log('event -> inStart');
//        })
//        .one('animsition.inEnd', function () {
            
//            console.log('event -> inEnd');
//        })
//        .one('animsition.outStart', function () {
//            console.log('event -> outStart');
//        })
//        .one('animsition.outEnd', function () {
//            console.log('event -> outEnd');
//        });
//});


/*Overlay 1 */

//$(document).ready(function () {
//    $(".animsition").animsition({
//        inClass: 'overlay-slide-in-left',
//        outClass: 'overlay-slide-in-top',
//        inDuration: 1500,
//        outDuration: 800,
//        linkElement: '.animsition-link',
//        // e.g. linkElement: 'a:not([target="_blank"]):not([href^="#"])'
//        loading: true,
//        loadingParentElement: 'body', //animsition wrapper element
//        loadingClass: 'animsition-loading',
//        loadingInner: '', // e.g '<img src="loading.svg" />'
//        timeout: false,
//        timeoutCountdown: 5000,
//        onLoadEvent: true,
//        browser: ['animation-duration', '-webkit-animation-duration'],
//        // "browser" option allows you to disable the "animsition" in case the css property in the array is not supported by your browser.
//        // The default setting is to disable the "animsition" in a browser that does not support "animation-duration".
//        overlay: true,
//        overlayClass: 'animsition-overlay-slide',
//        overlayParentElement: 'body',
//        transition: function (url) { window.location.href = url; }
//    });
//});

//fadeout left
//$(document).ready(function () {
//    $(".animsition").animsition({
//        inClass: 'fade-out-right',
//        outClass: 'fade-out-left',
//        inDuration: 1500,
//        outDuration: 800,
//        linkElement: '.animsition-link',
//        // e.g. linkElement: 'a:not([target="_blank"]):not([href^="#"])'
//        loading: true,
//        loadingParentElement: 'body', //animsition wrapper element
//        loadingClass: 'animsition-loading',
//        loadingInner: '', // e.g '<img src="loading.svg" />'
//        timeout: false,
//        timeoutCountdown: 5000,
//        onLoadEvent: true,
//        browser: ['animation-duration', '-webkit-animation-duration'],
//        // "browser" option allows you to disable the "animsition" in case the css property in the array is not supported by your browser.
//        // The default setting is to disable the "animsition" in a browser that does not support "animation-duration".
//        overlay: false,
//        overlayClass: 'animsition-overlay-slide',
//        overlayParentElement: 'body',
//        transition: function (url) { window.location.href = url; }
//    });
//});


// Options
//$(document).ready(function () {
//    $(".animsition").animsition({
//        inClass: 'zoom-in-sm',
//        outClass: 'zoom-out-sm',
//        inDuration: 500,
//        outDuration: 800,
//        linkElement: '.animsition-link',
//        // e.g. linkElement: 'a:not([target="_blank"]):not([href^="#"])'
//        loading: true,
//        loadingParentElement: 'body', //animsition wrapper element
//        loadingClass: 'animsition-loading',
//        loadingInner: '', // e.g '<img src="loading.svg" />'
//        timeout: false,
//        timeoutCountdown: 5000,
//        onLoadEvent: true,
//        browser: ['animation-duration', '-webkit-animation-duration'],
//        // "browser" option allows you to disable the "animsition" in case the css property in the array is not supported by your browser.
//        // The default setting is to disable the "animsition" in a browser that does not support "animation-duration".
//        overlay: false,
//        overlayClass: 'animsition-overlay-slide',
//        overlayParentElement: 'body',
//        transition: function (url) { window.location.href = url; }
//    });
//});
