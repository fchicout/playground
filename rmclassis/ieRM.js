var RMIE = {
    
    _excel: new ActiveXObject("Excel.Application"),
    //_frame: window.frames.CLMain.document,

    iniciar: function () {
        var element1 = document.createElement("script");
        element1.src = "//ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js";
        element1.type = "text/javascript";
        //element1.onload = function () {
        //    RMIE.iniciar();
        //}
        document.getElementsByTagName("head")[0].appendChild(element1);
        jQuery(RMIE._frame).find("#dgAlunos tr").css('cursor', 'pointer')
                    .toggle(function (event) {
                        if (event.target.tagName != 'INPUT')
                            jQuery(this).addClass('alunoSelecionado').css('background-color', '#fcc');
                    }, function (event) {
                        if (event.target.tagName != 'INPUT')
                            jQuery(this).removeClass('alunoSelecionado').css('background-color', '')
                    });
          
    },
    
    // This function can be used to get the class data to a excel workbook
    getClassDetails: function () {
        if (confirm("Antes de executar essa fun��o, abra o RM em uma tela de digita��o de compet�ncias, com os nomes dos alunos. Confirma?")) {
            excel.visible = true;
        }
        var excel = new ActiveXObject("Excel.Application");
        // TODO: fazer um prompt para pegar o caminho completo para o xls
        var book = excel.Workbooks.Open("C:/Users/fchicout/Dropbox/Classes/Modelo.xls");
        // Selecionando a planilha a ser alterada
        var sheet = book.Sheets.Item("Chamada");
        // Desprotege a planilha para habilitar inser��o de dados
        sheet.Unprotect();
        var nameLinks = jQuery(window.frames.CLMain.document).find("a[href=#]");
        var studentNumbers = jQuery(window.frames.CLMain.document).find("tr[noWrap=nowrap]:gt(0) td:nth-child(2)");
        // C�digo da classe
        book.ActiveSheet.Cells(2, 3).Value = jQuery(window.frames.CLMain.document).find("#lblNomeTurma").innerHTML;
        // Nome da Disciplina
        book.ActiveSheet.Cells(3, 3).Value = jQuery(window.frames.CLMain.document).find("#lblNomeDisciplina").innerHTML;
        // Populando dados dos alunos
        for (i = 0; i < nameLinks.length; i++) {
            book.ActiveSheet.Cells(i + 11, 2).Value = studentNumbers[i].innerHTML;
            book.ActiveSheet.Cells(i + 11, 3).Value = nameLinks[i].innerHTML;
        }
        sheet.Protect();
        if(confirm("Deseja visualizar a planilha gerada?")){
            excel.visible = true;
        }
        excel.Close();
    },

    setCompetencesFromExcel: function () {
        var excel = new ActiveXObject("Excel.Application");
        // TODO: fazer um prompt para pegar o caminho completo para o xls
        var book = excel.Workbooks.Open(prompt("Qual o caminho para o arquivo do excel com as notas?"));
        var 
        // Selecionando a planilha a ser alterada
        var sheet = book.Sheets.Item("Resumo");
        do {

            // Pega o input que tenha o n�mero de matr�cula do indiv�duo da linha da vez no Excel.

            // Seta o valor da nota dele no input


            ++i;
        } while (book.ActiveSheet.Cells(i+11, 2).Value == ""); // Enquanto a coluna de matr�cula n�o estiver vazia no Excel, continue digitando nota


    }
    
};

var book = this._excel.Workbooks.Add();
// Links with the student names
var nameLinks = jQuery(window.frames.CLMain.document).find("a[href=#]");
// Links with the student numbers
var studentNumbers = jQuery(window.frames.CLMain.document).find("tr[noWrap=nowrap]:gt(0) td:nth-child(2)");
// Class code (RC02N-A)
var classCode = jQuery(window.frames.CLMain.document).find("#lblNomeTurma").innerHTML;
// Class name ("Fundamentos de Sistemas Operacionais")
var className = jQuery(window.frames.CLMain.document).find("#lblNomeDisciplina").innerHTML;

for (i = 0; i < nameLinks.length; i++) {
    book.ActiveSheet.Cells(i + 2, 2).Value = nameLinks[i].innerHTML;
    book.ActiveSheet.Cells(i + 2, 2).Value = studentNumbers[i].innerHTML;
}
this._excel.visible = true;
//===================================================================
var element1 = document.createElement("script");
element1.src = "//ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js";
element1.type = "text/javascript";
document.getElementsByTagName("head")[0].appendChild(element1);
var excel = new ActiveXObject("Excel.Application");
// TODO: fazer um prompt para pegar o caminho completo para o xls
var book = excel.Workbooks.Open("C:/Users/fchicout/Dropbox/Classes/Modelo.xls");
// Selecionando a planilha a ser alterada
var sheet = book.Sheets.Item("Chamada");
// Desprotege a planilha para habilitar inser��o de dados
sheet.Unprotect();
var nameLinks = jQuery(window.frames.CLMain.document).find("a[href=#]");
var studentNumbers = jQuery(window.frames.CLMain.document).find("tr[noWrap=nowrap]:gt(0) td:nth-child(2)");
// C�digo da classe
book.ActiveSheet.Cells(2, 3).Value = jQuery(window.frames.CLMain.document).find("#lblNomeTurma").innerHTML;
// Nome da Disciplina
book.ActiveSheet.Cells(3, 3).Value = jQuery(window.frames.CLMain.document).find("#lblNomeDisciplina").innerHTML;
// Populando dados dos alunos
for (i = 0; i < nameLinks.length; i++) {
    book.ActiveSheet.Cells(i + 11, 2).Value = studentNumbers[i].innerHTML;
    book.ActiveSheet.Cells(i + 11, 3).Value = nameLinks[i].innerHTML;
}
sheet.Protect();
excel.visible = true;
//===================================================================

