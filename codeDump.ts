function timeBetweenDateAndTomorrowMorning(date) {


    // if (typeof date !== "date") {
    //   console.log("That wasn't a date!");
    //   return;
    // }

    /**
     * ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓ TODD'S TIME CALCULATION
     */
       console.log("Linked!")
       var midnite = new Date(date);
       midnite.setDate(midnite.getDate() + 1)
       midnite.setHours(0);
       midnite.setMinutes(0);
       midnite.setSeconds(0);
   
      //  var endOfDay = new Date();
      //  endOfDay.setHours(endHour);
      //  endOfDay.setMinutes(endMinute);
      //  endOfDay.setSeconds(0);
   
       var startOfDay = new Date(date);
       startOfDay.setDate(startOfDay.getDate() + 1);
       startOfDay.setHours(8);
       startOfDay.setMinutes(30);
       startOfDay.setSeconds(0);


      var midniteMilli = midnite.getTime();
      var dateMilli = date.getTime();
      var startMilli = startOfDay.getTime();


   
      //  var timeLeftTodayInMinutes = ((Math.abs(midnite.getTime() - date.getTime()) / 1000) / 60);
       var timeLeftTodayInHours = ((Math.abs(midniteMilli - dateMilli) / 1000) / 60) / 60;
   
      //  var timeLeftTomorrowInMinutes = ((Math.abs(startOfDay.getTime() - midnite.getTime()) / 1000) / 60);
       var timeLeftTomorrowInHours = ((Math.abs(startMilli - midniteMilli) / 1000) / 60) / 60;
   
       console.log(`Time between endOfDay and midnite
         in Hours: ${timeLeftTodayInHours}
   
         Time between midnite and startOfDay
         in Hours: ${timeLeftTomorrowInHours}
   
         Amount of time before tomorrow Morning
         ${timeLeftTodayInHours + timeLeftTomorrowInHours}
       `);

       return (timeLeftTodayInHours + timeLeftTomorrowInHours);
   
   
       /**
        * ↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑ TODD'S TIME CALCULATION
        */
}