var date = new Date();
//for simplicity, let's say the date is Wed, 9/22/21, 4:17PM
var adjustmentNumber = 7;
var adjustedDate = new Date(date);
adjustedDate.setHours(adjustedDate.getDate() + adjustmentNumber);

var sunday = {
  dayID: 0,
  startHour: 8,
  startMinute: 30,
  endHour: 17,
  endMinute: 30,
  workDay: 0,
}
var monday = {
  dayID: 1,
  startHour: 8,
  startMinute: 0,
  endHour: 17,
  endMinute: 0,
  workDay: 0,
}
var tuesday = {
  dayID: 2,
  startHour: 8,
  startMinute: 30,
  endHour: 17,
  endMinute: 30,
  workDay: 0,
}
var wednesday = {
  dayID: 3,
  startHour: 8,
  startMinute: 30,
  endHour: 17,
  endMinute: 30,
  workDay: 0,
}
var thursday = {
  dayID: 4,
  startHour: 8,
  startMinute: 0,
  endHour: 18,
  endMinute: 0,
  workDay: 0,
}
var friday = {
  dayID: 5,
  startHour: 8,
  startMinute: 30,
  endHour: 13,
  endMinute: 30,
  workDay: 0,
}
var saturday = {
  dayID: 6,
  startHour: 8,
  startMinute: 30,
  endHour: 17,
  endMinute: 30,
  workDay: 0,
}

function timeBetweenDateAndTomorrowMorning(date, adjustmentNumber, adjustedDate) { 

  var adjustedDayOfWeek = dayOfWeek(date); //returns a dayID (0-6) for the day of the week of the date object

  var dayTitle = titleDOW(adjustedDayOfWeek); //returns a day title based on the dayID of the adjustedDayOfWeek variable

  var startOfWorkDay = new Date(date); //adjusts start time of work day based on the day of the week
  startOfWorkDay.setHours(dayTitle.startHour);
  startOfWorkDay.setMinutes(dayTitle.startMinute);
  startOfWorkDay.setSeconds(0);

  var endOfWorkDay = new Date(date); //adjusts end time of work day based on the day of the week
  endOfWorkDay.setHours(dayTitle.endHour);
  endOfWorkDay.setMinutes(dayTitle.endMinute);
  endOfWorkDay.setSeconds(0);

  if (adjustedDate.getTime() < startOfWorkDay.getTime() || adjustedDate.getTime() > endOfWorkDay.getTime()) { //if adjustedDate falls outside of office hours, do this...
    //if in the evening, subtract end of office hours from date
    //if in the morning, subtract start of office hours from date
  }



  


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

function dayOfWeek(d) { //finds the day of the week
  var day = d.getDay();
  return day;
};

function titleDOW(d) { //returns the day of the week (refered to directly in another variable) based on the dayID index number
  if (d = 0) {
    return sunday;
  } else if (d = 1) {
    return monday;
  } else if (d = 2) {
    return tuesday;
  } else if (d = 3) {
    return wednesday;
  } else if (d = 4) {
    return thursday;
  } else if (d = 5) {
    return friday;
  } else if (d = 6) {
    return saturday;
  };
};

