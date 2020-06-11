function PickDueDate () {
/**
 *  add days to current date
 * 
 * @returns date that is X business days from today
*/ 
	var count = 0
	var dd, mm, yyyy
	
	// initialize dates
	var holidays = ["01/01", "07/04", "12/25"] // holidays that give error in CPLUS
	var endDate = new Date()

	if (endDate.getDate() < 15) {
		daysToAdd = 20
	} else {
		daysToAdd = 4
	}

	// set date ahead by 20 days (4 weeks excluding weekends)
	while (count < daysToAdd) {
		endDate.setDate(endDate.getDate() + 1)
		dd = String(endDate.getDate())
		mm = String(endDate.getMonth() + 1) // Jan is month 0
		
		for (each in holidays) {
			if (holidays[each].toString() == (mm+'/'+dd)) {
				continue
			}
		}

		// increment count if not sunday, saturday, or holiday
		if (endDate.getDay() != 0 && endDate.getDay() != 6) { 
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

	yyyy = String(endDate.getFullYear())
	return (mm + '/' + dd + '/' + yyyy).toString()
}

PickDueDate()