

export function filterItems_Of_VacationRequest(StartDate, EndDate, Status){
    // filter by start date, endDate and status
    var newVacationRequestsData
    if ((StartDate !== null) && (EndDate !== null) && (Status !== "")){
        console.log("test",1)
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
            new Date(item.dateDeDepart).setHours(0,0,0,0) <= new Date(StartDate).setHours(0,0,0,0) // compare start date with items start date 
            && new Date(item.EndDate).setHours(0,0,0,0) >= new Date(EndDate).setHours(0,0,0,0)  // compare end date with items end date 
            && item.RequestType === Status  // compare requestType and status of filter 
    )
    
    // filter by endDate and status
    }else if((EndDate !== null) && (Status !== "")){
        console.log("test",2)
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
            new Date(item.EndDate).setHours(0,0,0,0) <= new Date(EndDate).setHours(0,0,0,0) // compare end date with items end date 
            && item.RequestType === Status  // compare requestType and status of filter 
    )

    // filter by startDate and endDate
    }else if ((StartDate !== null) && (EndDate !== null)){
        console.log("test",3)
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
            new Date(item.dateDeDepart).setHours(0,0,0,0) === new Date(StartDate).setHours(0,0,0,0) // compare start date with items start date 
            && new Date(item.EndDate).setHours(0,0,0,0) === new Date(EndDate).setHours(0,0,0,0)  // compare end date with items end date 
    )

    // filter by startdate and status 
    }else if ((StartDate !== null) && (Status !== "")){
        console.log("test",4)
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
            new Date(item.dateDeDepart).setHours(0,0,0,0) <= new Date(StartDate).setHours(0,0,0,0) // compare start date with items start date 
            && item.RequestType === Status  // compare requestType and status of filter 
        )

    // if this filter is not in this conditions 
    }else{

        // filter by simple start date
        if (StartDate !== null){
            console.log("test",5)
            newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
                new Date(item.dateDeDepart).setHours(0,0,0,0) === new Date(StartDate).setHours(0,0,0,0) // compare start date with items start date 
            )
        }

        // filter by simple endDate
        if (EndDate !== null){
            console.log("test",6)
            newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
                new Date(item.EndDate).setHours(0,0,0,0) === new Date(EndDate).setHours(0,0,0,0) // compare end date with items end date 
            )
        }
        // filter by simple status of requests
        if (Status !== ""){
            console.log("test",7)
            newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
                item.RequestType === Status // compare requestType and status of filter 
            )
        }
    }
    return newVacationRequestsData
}