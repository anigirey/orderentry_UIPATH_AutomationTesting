function PickDueDate () {
/**
 *  add days to current date
 * 
 * @returns date that is X business days from today
*/ 
	var count = 0
	var dd, mm, yyyy
	
	// initialize dates
	var holidays = ["1/1", "7/4", "12/25"] // holidays that give error in CPLUS
	var endDate = new Date()
	yyyy = String(endDate.getFullYear())

	//calculate and add holidays to list
	holidays = holidays.concat(CalculateHolidays(yyyy))

	if (endDate.getDate() < 15) {
		daysToAdd = 20
	} else {
		daysToAdd = 4
	}

	// set date ahead by 20 days (4 weeks excluding weekends and holidays)
	var isHoliday = false
	while (count < daysToAdd) {
		isHoliday = false
		endDate.setDate(endDate.getDate() + 1)
		dd = String(endDate.getDate())
		mm = String(endDate.getMonth() + 1) // Jan is month 0
		
		for (each in holidays) {
			if (holidays[each].toString() == (mm+'/'+dd)) {
				isHoliday = true
				continue
			}
		}

		// increment count if not sunday, saturday, or holiday
		if (endDate.getDay() != 0 && endDate.getDay() != 6 && isHoliday == false) { 
			count++
		}
	}
	
	// format day and month
	if (dd.length < 2) {
		dd = '0' + dd
	}
	if (mm.length < 2) {
		mm = '0' + mm
	}

	return (mm + '/' + dd + '/' + yyyy)
}

function CalculateHolidays (year) {
	//month, week, day ; starting count from 0; -1 means last instance
	var holidays = [ 
		[4,-1,1], //"Memorial Day",
		[8,0,1],  //"Labor Day",
		[10,3,4], //"Thanksgiving Day"
	]

	var calculatedDates = []
	var month, week, day, firstDay
	var mm, dd

	for (each in holidays) {
		month = holidays[each][0]
		week = holidays[each][1]
		day = holidays[each][2]
		firstDay = 1;
		
		// if last instance of month, go to next month to work backwards
		if (week < 0) {
			month++;
			firstDay--;
		}
		
		var date = new Date(year, month, (week * 7) + firstDay);

		if (day < date.getDay()) {
			day += 7;
		}
		
		date.setDate(date.getDate() - date.getDay() + day);

		mm = String(date.getMonth() + 1)
		dd = String(date.getDate())

		calculatedDates.push(mm + "/" + dd)
	}
	return(calculatedDates)
}

PickDueDate()