function myFunc(callback) {

    var returnValue = false;
    var appCount = 0;

    $.ajax({
        type: "GET",
        url: "getBookedAppointmentCountByDoctorAndDate.php",
        data: dataString,
        success: function (response){
            appCount = response;
            //alert(appCount);
            if(appCount >= 6){
                returnValue = true;
            }
            callback(returnValue);
        }
    });
}



myFunc(function (value) {
    console.log(value);
    // do what ever you want with `value`
})
