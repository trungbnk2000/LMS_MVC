$(document).ready(function () {
    $('.btn-block').click(function () {
        var url = $('#myModal').data('url');
        $.get(url, function (data) {
            $('#myModal').html(data);
            $('#myModal').modal('show');
        });
    });
});