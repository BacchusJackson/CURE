$(function(){

    const $MH_Switch = $('#S01');
    const $PM_Switch = $('#S02'); 
    const $submitBtn = $('#submitBtn');

    console.log($PM_Switch);

    $('.datepicker').datepicker();
    $('select').formSelect();
    
    $('#clinicSwitch').on('click', function() {
        $MH_Switch.toggleClass('blue-text');
        $PM_Switch.toggleClass('red-text');
        $submitBtn.toggleClass('blue');
    });
    
});

