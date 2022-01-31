/* script para confirmação de botoes*/
var ladda = null;

$(document).on("click", "a[type='confirmar']", function (e) {
    var msg = $(this).data("mensagem");

    if (msg.length < 1)
        msg = "Deseja realmente excluir esse registro?";

    e.preventDefault();

    var href = $(this).context.href;

    swal({
        title: "Excluir",
        text: msg,
        icon: "error",
        buttons: true,
        dangerMode: true
    })
        .then((ok) => {
            if (ok) {
                window.location.replace(href);
            }
        });


});
function FormatDate(D) {
    try {
        var data = new Date(parseInt(D.substr(6)));
        var dia = data.getDate();
        if (dia.toString().length === 1)
            dia = "0" + dia;
        var mes = data.getMonth() + 1;
        if (mes.toString().length === 1)
            mes = "0" + mes;
        var ano = data.getFullYear();

    } catch (e) {
        return null;
    }
    return dia + "/" + mes + "/" + ano;

}

function DateFormat(date) {
    var day = (date.getDate().length == 1) ? "0" + date.getDate() : date.getDate();       // yields date
    var month = date.getMonth() + 1;    // yields month (add one as '.getMonth()' is zero indexed)
    var year = date.getFullYear();  // yields year
    var hour = "0" + date.getHours();     // yields hours
    var minute = "0" + date.getMinutes(); // yields minutes
    var second = "0" + date.getSeconds(); // yields seconds
    // After this construct a string with the above results as below
    return time = (day + "/" + month + "/" + year + " " + hour + ':' + minute + ':' + second);
}


function AguardeAlgunsEstantes(msg) {
    swal({
        title: '',
        text: typeof msg === 'undefined' ? "Aguarde alguns instantes." : msg,
        type: 'warning',
        showConfirmButton: false,
        allowOutsideClick: false,
        buttons: false
    });
}

$(document).ready(function () {
    /*==================
    select2
    ====================*/
    //$(document).ready(function () {
    //    $(".select2").select2();
    //});


    /*==================
    maskedinput
    ==================*/
    //$(function () {
    //    $(".mask-date").mask("99/99/9999");
    //    $(".mask-cpf").mask("999.999.999-99");
    //    $(".mask-cnpj").mask("99.999.999/9999-99");
    //    $(".mask-ciclo").mask("99");
    //    $(".mask-MesAno").mask("99/9999");


    //    $(".mask-cpf-cnpj").inputmask({
    //        mask: ["999.999.999-99", "99.999.999/9999-99"],
    //        keepStatic: true
    //    });


    //});

    //$(".mask-phone").mask("(99) 9999-9999?9").ready(function (event) {

    //    var target = (event.currentTarget) ? event.currentTarget : event.srcElement;

    //    if (target !== undefined) {

    //        var phone = target.value.replace(/\D/g, "");

    //        var element = $(target);

    //        element.unmask();

    //        if (phone.length > 10) {
    //            element.mask("(99) 99999-999?9");
    //        } else {
    //            element.mask("(99) 9999-9999?9");
    //        }
    //    }
    //});

    /*==================
    datepicker
    ==================*/
    $(".datepicker").datepicker({
        language: "pt-BR",
        format: "dd/mm/yyyy",
        autoclose: true,
        
    });

    $(".datepickerMonthYear").datepicker({
        language: "pt-BR",
        format: "mm/yyyy",
        minViewMode: "months",
        viewMode: "months",
        autoclose: true

    });

    /*==================
    icheck
    ==================*/
    $(".i-check").iCheck({
        checkboxClass: "icheckbox_square-blue",
        radioClass: "iradio_square-blue"
    });

    $(".i-check-all").on("ifChecked ifUnchecked", function (event) {
        if (event.type === "ifChecked") {
            $(".i-check").iCheck("check");
        } else {
            $(".i-check").iCheck("uncheck");
        }
    });


    /*==================
    datatable
    ==================*/
    var dataTable = $(".dataTable").DataTable({
        language: {
            "sEmptyTable": "Nenhum registro encontrado",
            "sInfo": "Mostrando de _START_ até _END_ de _TOTAL_ registros",
            "sInfoEmpty": "Mostrando 0 até 0 de 0 registros",
            "sInfoFiltered": "(Filtrados de _MAX_ registros)",
            "sInfoPostFix": "",
            "sInfoThousands": ".",
            "sLengthMenu": "_MENU_ resultados por página",
            "sLoadingRecords": "Carregando...",
            "sProcessing": "Processando...",
            "sZeroRecords": "Nenhum registro encontrado",
            "sSearch": "Pesquisar",
            "oPaginate": {
                "sNext": "Próximo",
                "sPrevious": "Anterior",
                "sFirst": "Primeiro",
                "sLast": "Último"
            },
            "oAria": {
                "sSortAscending": ": Ordenar colunas de forma ascendente",
                "sSortDescending": ": Ordenar colunas de forma descendente"
            }
        }
    });




   



    $(document).on("ifChecked", ".i-check-all", null, function () {

        var listaCheckBox = $(".i-check", dataTable.cells().nodes());

        for (var i = 0; i < listaCheckBox.length; i++) {
            $(listaCheckBox[i]).iCheck("check");
        }

    });

    $(document).on("ifUnchecked",
        ".i-check-all",
        null,
        function () {

            var listaCheckBox = $(".i-check", dataTable.cells().nodes());

            for (var i = 0; i < listaCheckBox.length; i++) {
                $(listaCheckBox[i]).iCheck("uncheck");
            }

        });

});

var dataTableConfig = {
    language: {
        "sEmptyTable": "Nenhum registro encontrado",
        "sInfo": "Mostrando de _START_ até _END_ de _TOTAL_ registros",
        "sInfoEmpty": "Mostrando 0 até 0 de 0 registros",
        "sInfoFiltered": "(Filtrados de _MAX_ registros)",
        "sInfoPostFix": "",
        "sInfoThousands": ".",
        "sLengthMenu": "_MENU_ resultados por página",
        "sLoadingRecords": "Carregando...",
        "sProcessing": "Processando...",
        "sZeroRecords": "Nenhum registro encontrado",
        "sSearch": "Pesquisar",
        "oPaginate": {
            "sNext": "Próximo",
            "sPrevious": "Anterior",
            "sFirst": "Primeiro",
            "sLast": "Último"
        },
        "oAria": {
            "sSortAscending": ": Ordenar colunas de forma ascendente",
            "sSortDescending": ": Ordenar colunas de forma descendente"
        }
    }
}