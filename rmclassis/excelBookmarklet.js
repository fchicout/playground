function RMIE() {
    this._excel = new ActiveXObject("Excel.Application");
    var element1 = document.createElement("script");
    element1.src = "//ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js";
    element1.type = "text/javascript";
    element1.onload = function () {
        document.getElementsByTagName("head")[0].appendChild(element1);
        jQuery(RMIE._frame).find("#dgAlunos tr").css('cursor', 'pointer')
                    .toggle(function (event) {
                        if (event.target.tagName != 'INPUT')
                            jQuery(this).addClass('alunoSelecionado').css('background-color', '#fcc');
                    }, function (event) {
                        if (event.target.tagName != 'INPUT')
                            jQuery(this).removeClass('alunoSelecionado').css('background-color', '')
                    });

    }

    this.getClassDetails = function () {
        if (confirm("Antes de executar essa função, abra o RM em uma tela de digitação de competências, com os nomes dos alunos. Confirma?")) {
            excel.visible = true;
        }
        var excel = new ActiveXObject("Excel.Application");
        // TODO: fazer um prompt para pegar o caminho completo para o xls
        var book = excel.Workbooks.Open("C:/Users/fchicout/Dropbox/Classes/Modelo.xls");
        // Selecionando a planilha a ser alterada
        var sheet = book.Sheets.Item("Chamada");
        // Desprotege a planilha para habilitar inserção de dados
        sheet.Unprotect();
        var nameLinks = jQuery(window.frames.CLMain.document).find("a[href=#]");
        var studentNumbers = jQuery(window.frames.CLMain.document).find("tr[noWrap=nowrap]:gt(0) td:nth-child(2)");
        // Código da classe
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
    }
    this.setCompetencesFromExcel = function () {
        var excel = new ActiveXObject("Excel.Application");
        // TODO: fazer um prompt para pegar o caminho completo para o xls
        var book = excel.Workbooks.Open(prompt("Qual o caminho para o arquivo do excel com as notas?"));
        // Etapa de nota (C1, R4, etc...)
        var step = prompt("Qual a competência a trazer?");
        var stepCol = parseInt(step) + 18; // TODO: Precisa ser ajustada
        // Selecionando a planilha a ser alterada
        var sheet = book.Sheets.Item("Resumo");
        do {
            var studentNumber = book.ActiveSheet.Cells(i + 11, 2).Value;
            // Pega o input que tenha o número de matrícula do indivíduo da linha da vez no Excel e Seta o valor da nota dele no input
            jQuery(window.frames.CLMain.document).find("tr:contains(" + studentNumber + ")").find("input").val(book.ActiveSheet.Cells(i + 11, stepCol).Value);
            
            ++i;
        } while (book.ActiveSheet.Cells(i+11, 2).Value == ""); // Enquanto a coluna de matrícula não estiver vazia no Excel, o script continua digitando nota.


    }
}