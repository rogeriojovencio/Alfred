$(document).ready(function () {
    var table = $('#tbLotes').DataTable();
    table.destroy();

    $('#tbLotes').DataTable({
        order: [[1, "desc"]],
        paging: false,
        language: dataTableConfig.language
    });

});