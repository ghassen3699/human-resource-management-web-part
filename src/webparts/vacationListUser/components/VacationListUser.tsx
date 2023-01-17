import * as React from 'react';
import styles from './VacationListUser.module.scss';
import { IVacationListUserProps } from './IVacationListUserProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {
  Fabric,
  DefaultButton,
  Slider,
  Panel,
  PanelType,
  loadTheme
} from "office-ui-fabric-react";
import { DatePicker, IDatePickerStrings, IDatePickerStyles, IDatePickerStyleProps } from 'office-ui-fabric-react/lib/DatePicker';
import { Dropdown, IDropdownOption, IDropdownProps, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'VacationListUserWebPartStrings';
import { MSGraphClient, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';



export default class VacationListUser extends React.Component<IVacationListUserProps, {}> {
  // state of the webpart 
  public state = {
    vacationRequestsData : [],
    VacationListUserFilter : [],
    dateDebut: null,
    dateFin: null ,
    filterStatus: "",
    status: "Tous",


    // // Pagination params 
    // itemCount : 0,     // Lenght of vacation list user
    // pageSize: 7,       // size of page in vacation list table
    // currentPage: 1,    // current page in the pagination
    // pagesCount: 0,     // number of pages in pagination

    currentPage: 1,
    recordsPerPage: 7,

  }

  // Get users data from graphAPI
  public getUsers = async() => {
    var userData = []
    await this.props.context.msGraphClientFactory
    .getClient()
    .then((MSGraphClient:MSGraphClient) => {
      MSGraphClient
        .api("users")
        .version('v1.0')
        .select("displayName,mail,ID,userPrincipalName")
        .get((err, res)=> {
          if (err){
            console.log(err)
          }
          let currentUser = res.value.filter(user =>  this.props.context.pageContext.legacyPageContext["userPrincipalName"].toLowerCase() == user.userPrincipalName.toLowerCase())
          userData.push(currentUser[0])
        })
    })
  }


  // get vacation data of user
  public getVacationData = async() => {
    let userID = (await Web(this.props.url).currentUser()).Id
    const vacationRequests = await Web(this.props.url).lists.getByTitle('vacationRequest').items.filter("AuthorId eq " + userID)();
    this.setState({
      vacationRequestsData: vacationRequests, 
      VacationListUserFilter: vacationRequests,
    })
  }


  // change the format of vacation request start date and end date
  public formatDate = (date) => {
    date = new Date(date)
    let newDateFormat = date.getDate() + "/" + (date.getMonth() + 1) + "/" + date.getFullYear()
    return newDateFormat
  }


  // cancel the vacation request
  public cancelVacationRequest = async(ID) => {
    let cancel = await Web(this.props.url).lists.getByTitle('vacationRequest').items.getById(ID).update({RequestType:"Annulé"})
    window.location.reload()
  }


  // filter vacation requests by Start date, End date, and status
  public filterVacationItems = (StartDate, EndDate, Status) => {
    var newVacationRequestsData


    // filter by start date, endDate and status
    if ((StartDate !== null) && (EndDate !== null) && (Status !== "")){
      if (Status !== 'Tous'){
        // console.log("test",11)   
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
          item.RequestType === Status &&   // filter by the status of items and status selected in filter form
          new Date(StartDate).setHours(0,0,0,0) <= new Date(item.dateDeDepart).setHours(0,0,0,0) &&  // filter by startDate of items and startDate selected in filter 
          new Date(EndDate).setHours(0,0,0,0) >= new Date(item.EndDate).setHours(0,0,0,0)   // filter by endDate of items and endDate selected
        )
      }else {
        // console.log("test",1)
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
          new Date(item.dateDeDepart).setHours(0,0,0,0) >= new Date(StartDate).setHours(0,0,0,0) // compare start date with items start date 
          && new Date(item.EndDate).setHours(0,0,0,0) <= new Date(EndDate).setHours(0,0,0,0)  // compare end date with items end date 
        )
      }
      
    
    // filter by endDate and status
    }else if((EndDate !== null) && (Status !== "")){
      // console.log("test",2)
      if (Status !== 'Tous'){
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
          new Date(item.EndDate).setHours(0,0,0,0) === new Date(EndDate).setHours(0,0,0,0) // compare end date with items end date 
          && item.RequestType === Status  // compare requestType and status of filter 
        )
      }else {
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
          new Date(item.EndDate).setHours(0,0,0,0) === new Date(EndDate).setHours(0,0,0,0) // compare end date with items end date 
        )
      }


    // filter by startDate and endDate
    }else if ((StartDate !== null) && (EndDate !== null)){
      // console.log("test",3)
      newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
        new Date(item.dateDeDepart).setHours(0,0,0,0) === new Date(StartDate).setHours(0,0,0,0) // compare start date with items start date 
        && new Date(item.EndDate).setHours(0,0,0,0) === new Date(EndDate).setHours(0,0,0,0)  // compare end date with items end date 
      )

    // filter by startdate and status 
    }else if ((StartDate !== null) && (Status !== "")){
      // console.log("test",4)
      if (Status !== 'Tous'){
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
          new Date(item.dateDeDepart).setHours(0,0,0,0) === new Date(StartDate).setHours(0,0,0,0) // compare start date with items start date 
          && item.RequestType === Status  // compare requestType and status of filter 
        )
      }else{
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
          new Date(item.dateDeDepart).setHours(0,0,0,0) === new Date(StartDate).setHours(0,0,0,0) // compare start date with items start date 
        )
      }
      

    // if this filter is not in this conditions 
    }else{

      // filter by simple start date
      if (StartDate !== null){
        // console.log("test",5)
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
          new Date(item.dateDeDepart).setHours(0,0,0,0) === new Date(StartDate).setHours(0,0,0,0) // compare start date with items start date 
        )
      }

      // filter by simple endDate
      if (EndDate !== null){
        // console.log("test",6)
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
          new Date(item.EndDate).setHours(0,0,0,0) === new Date(EndDate).setHours(0,0,0,0) // compare end date with items end date 
        )
      }
      // filter by simple status of requests
      if ((Status !== "")&&(Status !== "Tous")){
        // console.log("test",7)
        newVacationRequestsData = this.state.vacationRequestsData.filter(item => 
          item.RequestType === Status // compare requestType and status of filter 
        )
      }else {
        newVacationRequestsData = this.state.vacationRequestsData
      }
    }
    this.setState({VacationListUserFilter:newVacationRequestsData})  
  }

  // send item to update page
  public sendUpdatedItem = (itemID) => {
    window.location.href = this.props.context.pageContext.web.absoluteUrl + "/SitePages/Update-Vaca.aspx?itemId="+itemID
  }


  // // Increment pagination in vacation list 
  // public paginationHandleClickIncrement = (pageNumber, itemCount, pagesCount) => {
  //   const startIndex = (pageNumber - 1) * this.state.pageSize
  //   console.log(startIndex)
  //   const data = this.state.VacationListUserFilter.slice(startIndex,this.state.pageSize)
  //   this.setState({currentPage:pageNumber+1})
  //   console.log('current page',this.state.currentPage)
  //   console.log('data',data)
  // }

  // // decrement pagination in vacation list
  // public paginationHandleClickDecrement = (pageNumber, itemCount, pagesCount) => {
  //   if (this.state.currentPage !== 0){
  //     this.setState({currentPage:pageNumber-1})
  //     console.log(this.state.currentPage)
  //   }

  // }


  public decrementPagination = (currentPage, indexOfLastRecord, indexOfFirstRecord) => {
    this.setState({currentPage:currentPage - 1})
  }

  public incrementPagination = (currentPage, indexOfLastRecord, indexOfFirstRecord) => {
    var data = this.state.VacationListUserFilter.slice(indexOfFirstRecord, indexOfLastRecord)
    console.log(data)
    this.setState({
      currentPage:currentPage + 1,
    })
  }


  // update Get Users Stat when initialise page
  componentDidMount(): void {
    this.getUsers();
    this.getVacationData();
  }


  public render(): React.ReactElement<IVacationListUserProps> {

    // Style of Dropdown
    const dropdownStyles: Partial<IDropdownStyles> = {
      dropdown: { width: 160 },
    };

     

    // date picker info
    const DatePickerStrings: IDatePickerStrings = {
      months: ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Aout', 'Septembre', 'Octobre', 'Novembre', 'Décembre'],
      shortMonths: ['Jan', 'Feb', 'Mar', 'Avr', 'Mai', 'Jun', 'Jul', 'Aou', 'Sep', 'Oct', 'Nov', 'Dec'],
      days: ['Diamanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'],
      shortDays: ['DI', 'LU', 'MA', 'ME', 'JE', 'VE', 'SA'],
      goToToday: "Aller à aujourd'hui",
      prevMonthAriaLabel: 'Aller au mois précédent',
      nextMonthAriaLabel: 'Aller au mois prochain',
      prevYearAriaLabel: "Aller à l'année précédente",
      nextYearAriaLabel: "Aller à l'année prochaine",
      invalidInputErrorMessage: 'Invalid date format.'
    };

    const indexOfLastRecord = this.state.currentPage * this.state.recordsPerPage
    const indexOfFirstRecord = indexOfLastRecord - this.state.recordsPerPage;
    

    return (
        <div className={styles.vacationListUser}>
          <div className={styles.title}><strong>Filtres</strong></div>
          <div className={styles.filters}>
            <label className={styles.title}>Statut :  &nbsp;</label>
            <Dropdown
                styles={dropdownStyles}
                options={[
                  {key: "Tous",text: "Tous"},
                  {key: "en cours",text: "en cours"},
                  {key: "Annulé", text: "Annulé"}, 
                  {key: "Validé", text: "Validé"}, 
                ]}
                placeholder="Select an option"  
                defaultSelectedKey={this.state.status}
                onChanged={(value) => this.setState({status:value['key']})}
              />

            <label className={styles.title}>Date de départ : </label>
            <DatePicker
                  className={styles.startDate}
                  allowTextInput={false}
                  strings={DatePickerStrings}
                  onSelectDate={(selectedStartDate) =>  { this.setState({dateDebut:selectedStartDate}) }}
                  ariaLabel="Select a date"
                  value={this.state.dateDebut}
                />

            <label className={styles.title}>Date de fin : </label>
            <DatePicker
                  className={styles.startDate}
                  allowTextInput={false}
                  strings={DatePickerStrings}
                  onSelectDate={(selectedEndDate) => { this.setState({dateFin:selectedEndDate}) }}
                  ariaLabel="Select a date"
                  value={this.state.dateFin}
                />
            <div className={styles.title} id="SoldeDeConges">Solde de conges : 15</div>
            <button className={styles.btnRef} id={'refreshbutton'} onClick={() => this.filterVacationItems(this.state.dateDebut,this.state.dateFin,this.state.status)}>Filtrer</button>
          </div>

          <div id="spListContainer" > 
            <table style={{borderCollapse:"collapse", width:'100%'}}>
              <tr>
                <th className={styles.textCenter}>#</th> 
                <th>Motif d'absence</th>
                <th>Date de début</th>
                <th>Date de fin</th>
                <th>Jours de congés</th>
                <th>Statut d'approbation</th>
                <th></th>
                <th></th>
              </tr>
              {this.state.VacationListUserFilter.map(item => 
                <tr id={item.ID}>
                  <td></td> 
                  <td>{item.ctgVacation}</td> 
                  <td>{this.formatDate(item.dateDeDepart)}</td> 
                  <td>{this.formatDate(item.EndDate)}</td> 
                  <td>{item.NrbDays}</td> 
                  {/* If request type is "En cours" */}
                  {item.RequestType === "en cours" && <td>
                    <div className={styles.cercleBleu}></div> {item.RequestType}  
                  </td>}

                  {/* If request type is "Annuler" */}
                  {item.RequestType === "Annulé" && <td>
                    <div className={styles.cercleVert}></div> {item.RequestType}  
                  </td>}

                  {/* If request type is "Validé" */}
                  {item.RequestType === "Validé" && <td>
                    <div className={styles.cercleRouge}></div> {item.RequestType}  
                  </td>}

                  {item.RequestType !== 'Annulé' && 
                  
                    <>
                      {/* Update vacation request */}
                      <td>  
                        <a href="#">
                          <span className={styles.btnApprove} onClick={() => this.sendUpdatedItem(item.ID)}>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className={"bi bi-pencil-square"} viewBox="0 0 16 16">
                              <path d="M15.502 1.94a.5.5 0 0 1 0 .706L14.459 3.69l-2-2L13.502.646a.5.5 0 0 1 .707 0l1.293 1.293zm-1.75 2.456-2-2L4.939 9.21a.5.5 0 0 0-.121.196l-.805 2.414a.25.25 0 0 0 .316.316l2.414-.805a.5.5 0 0 0 .196-.12l6.813-6.814z"/>
                              <path fill-rule="evenodd" d="M1 13.5A1.5 1.5 0 0 0 2.5 15h11a1.5 1.5 0 0 0 1.5-1.5v-6a.5.5 0 0 0-1 0v6a.5.5 0 0 1-.5.5h-11a.5.5 0 0 1-.5-.5v-11a.5.5 0 0 1 .5-.5H9a.5.5 0 0 0 0-1H2.5A1.5 1.5 0 0 0 1 2.5v11z"/>
                            </svg>
                          </span>
                        </a>
                      </td>

                      {/* Cancel Vacation request */}
                      <td style={{cursor:'pointer'}} onClick={()=>this.cancelVacationRequest(item.ID)}>  
                          <span className={styles.btnRefuse}><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" className={"bi bi-x-square"} viewBox="0 0 16 16">
                            <path d="M14 1a1 1 0 0 1 1 1v12a1 1 0 0 1-1 1H2a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1h12zM2 0a2 2 0 0 0-2 2v12a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V2a2 2 0 0 0-2-2H2z"/>
                            <path d="M4.646 4.646a.5.5 0 0 1 .708 0L8 7.293l2.646-2.647a.5.5 0 0 1 .708.708L8.707 8l2.647 2.646a.5.5 0 0 1-.708.708L8 8.707l-2.646 2.647a.5.5 0 0 1-.708-.708L7.293 8 4.646 5.354a.5.5 0 0 1 0-.708z"/>
                            </svg>
                          </span>
                      </td>
                    </>
                  }
                  {item.RequestType === 'Annulé' && <><td></td><td></td></>}
                  
                </tr>
              )}
            </table>
          
          </div>


          


          <div className={styles.paginations}>
            <button className={styles.pagination} onClick={() => this.decrementPagination(this.state.currentPage, indexOfLastRecord, indexOfFirstRecord)}>Prev</button> 
            <span id="page"></span>
            <button className={styles.pagination} onClick={() => this.incrementPagination(this.state.currentPage, indexOfLastRecord, indexOfFirstRecord)}>Next</button>
          </div>
          
          {/* <div className={styles.cercleBleu}></div>En cours */}
          {/* <div id="myModal" className={styles.modal}></div> */}
          
        </div>
    );
  }
}
