var APP = {
    _frame: window.frames.CLMain.document,

    iniciar: function () {
        jQuery(APP._frame).find("#dgAlunos tr").css('cursor', 'pointer')
                    .toggle(function (event) {
                        if (event.target.tagName != 'INPUT')
                            jQuery(this).addClass('alunoSelecionado').css('background-color', '#fcc');
                    }, function (event) {
                        if (event.target.tagName != 'INPUT')
                            jQuery(this).removeClass('alunoSelecionado').css('background-color', '')
                    });
    },

    aplicarATodos: function () {
        var nota = prompt('Digite a nota');
        if (!nota) return false;
        jQuery(window.frames.CLMain.document)
                    .find('input[type=text][name$=Nota]')
                                .val(nota);

    },

    aplicarAMarcados: function () {
        var nota = prompt('Digite a nota');
        if (!nota) return false;

        jQuery(window.frames.CLMain.document)
                    .find("#dgAlunos tr.alunoSelecionado")
                                .find("input[type=text][name$=Nota]")
                                            .val(nota)
                                                        .end().removeClass('alunoSelecionado');
    },

    getStudentNames: function () {
        var studentNames = new Array();
        jQuery(window.frames.CLMain.document).find("a[href=#]");
        for (i = 0; i < links.length; i++) {
            studentNames.push(links[i].innerHTML);
        }
    }
};
var element1 = document.createElement("script");
element1.src = "//ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js";
element1.type = "text/javascript";
element1.onload = function () {
    APP.iniciar();
}
document.getElementsByTagName("head")[0].appendChild(element1);
