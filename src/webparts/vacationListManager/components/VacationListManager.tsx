import * as React from 'react';
import styles from './VacationListManager.module.scss';
import { IVacationListManagerProps } from './IVacationListManagerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DatePicker, Dropdown, IDatePickerStrings, IDropdownStyles } from 'office-ui-fabric-react';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { IFrameDialog } from "@pnp/spfx-controls-react/lib/IFrameDialog";
import { MSGraphClient, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';


export default class VacationListManager extends React.Component<IVacationListManagerProps, {}> {
  public state = {
    // data of vacation request 
    vacationRequestData : [],

    // data of vacation request after filter
    vacationRequestDataFilter : [],

    // option of Employees in Sharepoint site
    optionsEmployees : [],

    // PopUp information
    popUpUsername: "",
    popUpMotif: "",
    popUpDA: null,
    popUpDE: null,
    NumberOfDays: 0,
    Solde: "",
    popUpApprobation: "",
    popUpStatus : "",
    popUpComment : "",

    // filter options 
    EmployesFilter : "",
    StatusFilter : "",
    DateDebutFilter: null,
    DateFinFilter: null,
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
          userData.push({key:"Tous",text:"Tous"})
          res.value.map(user => {
            userData.push({
              key:user.userPrincipalName,
              text:user.displayName,
            })
          })
        })
    })
    this.setState({optionsEmployees:userData})
  }


  // Get user information with ID
  public getUserName = async(ID_User) => {
    var user = await Web(this.props.url).getUserById(ID_User)();
    return user.Title
  }


  // Get all data of vacation request 
  public vacationRequestData = async () => {
    var result = []
    var data = await Web(this.props.url).lists.getByTitle('vacationRequest').items.getAll();
    for (let i = 0; i < data.length; i++) {
      
      var user = await this.getUserName(data[i].AuthorId)
      console.log(user)
      result.push({
        Title: user,
        MotifAbsence: data[i].ctgVacation,
        DA: data[i].dateDeDepart,
        DE: data[i].EndDate,
        RequestType: data[i].RequestType,
        numbreOfDays: data[i].NrbDays,
        soldeOfDays: data[i].RemainingDays,
        Comment: data[i].Comment,
        ID: data[i].ID
      })
      
    }
    this.setState({
      vacationRequestData:result,
      vacationRequestDataFilter:result,
    })
  }

  // change the format of vacation request start date and end date
  public formatDate = (date) => {
    date = new Date(date)
    let newDateFormat = date.getDate() + "/" + (date.getMonth() + 1) + "/" + date.getFullYear()
    return newDateFormat
  }


  // popUp of details item in the table 
  public OpenPopUp = (userName, Motif, DA, DE, NumberOfDays, Solde, Approbation, Status, Commentaire) => {
    this.setState({
      popUpUsername: userName,
      popUpMotif: Motif,
      popUpDA: DA,
      popUpDE: DE,
      NumberOfDays: NumberOfDays,
      Solde: Solde,
      popUpApprobation: Approbation,
      popUpStatus: Status,
      popUpComment: Commentaire
    })
  }


  // filter items 
  public filterItems = async (dateDebut, dateFin, employee, status) => {

    var vacationRequestData = this.state.vacationRequestData 
    var dataAfterFilter

    if (employee !== "" && status !== "" && dateDebut === null && dateFin === null){         // Filter by Employee and status
      console.log(1)
      if (status !== "Tous"){
        dataAfterFilter = vacationRequestData.filter(item => item.Title === employee && item.RequestType === status)
      }else {
        dataAfterFilter = vacationRequestData.filter(item => item.Title === employee)

      }


    }else if (employee !== "" && dateDebut !== null && status === "" && dateFin === null){   // filter by Employee and date debut 
      console.log(2)
      dataAfterFilter = vacationRequestData.filter(item => item.Title === employee && new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0))

    }else if (employee !== "" && dateFin !== null && status === "" && dateDebut === null){   // filter by employee and date fin
      console.log(3)
      dataAfterFilter = vacationRequestData.filter(item => item.Title === employee && new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0))
      


    }else if (employee !== "" && status !== "" && dateDebut !== null && dateFin === null){   // filter by employee, status and date debut 
      console.log(4)
      if (status !== "Tous") {
        dataAfterFilter = vacationRequestData.filter(item => item.Title === employee && new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0) && item.RequestType === status)

      }else {
        dataAfterFilter = vacationRequestData.filter(item => item.Title === employee && new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0))

      }


    }else if (employee !== "" && status !== "" && dateFin !== null && dateDebut === null){   // filter by employee, status and date fin
      console.log(5)
      if (status !== "Tous") {
        dataAfterFilter = vacationRequestData.filter(item => item.Title === employee && new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0) && item.RequestType === status)

      }else {
        dataAfterFilter = vacationRequestData.filter(item => item.Title === employee && new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0))

      }


    }else if (employee !== "" && dateDebut !== null && dateFin !== null && status === ""){   // filter by employee, date debut and datefin
      console.log(6)
      dataAfterFilter = vacationRequestData.filter(item => item.Title === employee && new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0) && new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0))

    }else if (dateDebut !== null && dateFin !== null && status === "" && employee === ""){   // filter by date debut et date fin 
      console.log(7)
      dataAfterFilter = vacationRequestData.filter(item => new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0) && new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0))
      


    }else if (status !== "" && dateDebut !== null && dateFin === null && employee === ""){   // filter by status et date debut
      console.log(8)
      if (status !== "Tous") {
        dataAfterFilter = vacationRequestData.filter(item => new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0) && item.RequestType === status)

      }else {
        dataAfterFilter = vacationRequestData.filter(item => new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0))

      }


    }else if (status !== "" && dateFin !== null && dateDebut === null && employee === ""){   // filter by status et date fin
      console.log(9)
      if (status !== "Tous") {
        dataAfterFilter = vacationRequestData.filter(item => new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0) && item.RequestType === status)

      }else {
        dataAfterFilter = vacationRequestData.filter(item => new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0))

      }


    }else if (status !== "" && dateDebut !== null && dateFin !== null && employee === ""){   // filter by status, date debut et date fin
      console.log(10) 
      if (status !== "Tous") {
        dataAfterFilter = vacationRequestData.filter(item => item.RequestType === status && new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0) && new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0))

      }else {
        dataAfterFilter = vacationRequestData.filter(item => new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0) && new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0))

      }


    }else if (employee !== "" && status === "" && dateDebut === null && dateFin === null){   // filter by employee
      console.log(11)
      dataAfterFilter = vacationRequestData.filter(item => item.Title === employee)

    }else if (status !== "" && employee === "" && dateDebut === null && dateFin === null){   // filter by status 
      console.log(12)
      if (status !== "Tous") {
        dataAfterFilter = vacationRequestData.filter(item => item.RequestType === status)

      }else {
        dataAfterFilter = vacationRequestData

      }


    }else if (dateDebut !== null && employee === "" && dateFin === null && status === ""){   // filter by date debut 
      console.log(13)
      dataAfterFilter = vacationRequestData.filter(item =>  new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0))

    }else if (dateFin !== null && employee === "" && dateDebut === null && status === ""){  // filter by date fin
      console.log(14)
      dataAfterFilter = vacationRequestData.filter(item =>  new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0))

    }else if (dateFin !== null && employee !== "" && dateDebut !== null && status !== ""){                                                                                     // filter by employee, date debut, date fin et status
      console.log(15)
      if (status !== "Tous") {
        dataAfterFilter = vacationRequestData.filter(item => item.RequestType === status && new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0) && new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0) && item.Title === employee)

      }else {
        dataAfterFilter = vacationRequestData.filter(item => new Date(item.DE).setHours(0,0,0,0) <= new Date(dateFin).setHours(0,0,0,0) && new Date(item.DA).setHours(0,0,0,0) >= new Date(dateDebut).setHours(0,0,0,0) && item.Title === employee)

      }

    }

    this.setState({vacationRequestDataFilter:dataAfterFilter})
  }


  



  // update Get Users Stat when initialise page
  componentDidMount(): void {
    this.vacationRequestData();
    this.getUsers()
  }

  
  public render(): React.ReactElement<IVacationListManagerProps> {
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


    
    return (
      <div className={styles.vacationListUser}>
          <div className={styles.title}><strong>Filtres</strong></div>
          <div className={styles.filters}>

            <label className={styles.title}>Employés :  &nbsp;</label>
            <Dropdown
                styles={dropdownStyles}
                options={this.state.optionsEmployees}
                placeholder="Select an option"  
                // defaultSelectedKey={this.state.status}
                onChanged={(value) => this.setState({EmployesFilter:value.text})}
              />

            <label className={styles.title}>Status :  &nbsp;</label>
            <Dropdown
                styles={dropdownStyles}
                options={[
                  {key: "Tous",text: "Tous"},
                  {key: "en cours",text: "en cours"},
                  {key: "Annulé", text: "Annulé"}, 
                  {key: "Validé", text: "Validé"}, 
                ]}
                placeholder="Select an option"  
                // defaultSelectedKey={this.state.status}
                onChanged={(value) => this.setState({StatusFilter:value.key})} 
              />
              
            <label className={styles.title}>Date de départ : </label>
            <DatePicker
                  className={styles.startDate}
                  allowTextInput={false}
                  strings={DatePickerStrings}
                  onSelectDate={(selectedStartDate) =>  { this.setState({DateDebutFilter:selectedStartDate}) }}
                  ariaLabel="Select a date"
                  value={this.state.DateDebutFilter}
                />

            <label className={styles.title}>Date de fin : </label>
            <DatePicker
                  className={styles.startDate}
                  allowTextInput={false}
                  strings={DatePickerStrings}
                  onSelectDate={(selectedEndDate) => { this.setState({DateFinFilter:selectedEndDate}) }}
                  ariaLabel="Select a date"
                  value={this.state.DateFinFilter}
                />
            <button className={styles.btnRef} id={'refreshbutton'} onClick={() => this.filterItems(this.state.DateDebutFilter, this.state.DateFinFilter, this.state.EmployesFilter, this.state.StatusFilter)}>Filtrer</button>
          </div>

          <div id="spListContainer"> 
            <table style={{borderCollapse:"collapse", width:'100%'}}>
              <tr>
                <th className={styles.textCenter}>#</th> 
                <th>Nom et Prenom</th>
                <th>Motif d'absence</th>
                <th>Date de debut</th>
                <th>Date de fin</th>
                <th>Status d'approbation</th>
                <th>Détails</th>
              </tr>
              {/* A completer */}
              {this.state.vacationRequestDataFilter.map(vacationRequest => 
                <tr>
                  <td></td>
                  <td>{vacationRequest.Title}</td> 
                  <td>{vacationRequest.MotifAbsence}</td> 
                  <td>{this.formatDate(vacationRequest.DA)}</td> 
                  <td>{this.formatDate(vacationRequest.DE)}</td> 
                  <td>{vacationRequest.RequestType}</td> 
                  <td>
                    <a href="#popup1" onClick={()=> this.OpenPopUp(vacationRequest.Title, vacationRequest.MotifAbsence, vacationRequest.DA, vacationRequest.DE, vacationRequest.numbreOfDays, vacationRequest.soldeOfDays, vacationRequest.RequestType, vacationRequest.RequestType, vacationRequest.Comment)}>
                      <span className={styles.icon}>
                        <svg version="1.1" id="Capa_1"
                          xmlns="http://www.w3.org/2000/svg"
                          x="0px" y="0px" viewBox="0 0 512 512" style={{height:16, width:16}}>
                          <g>
                            <g>
                              <path d="M414.007,148.75c5.522,0,10-4.477,10-10V30c0-16.542-13.458-30-30-30h-364c-16.542,0-30,13.458-30,30v452
                                c0,16.542,13.458,30,30,30h364c16.542,0,30-13.458,30-30v-73.672c0-5.523-4.478-10-10-10c-5.522,0-10,4.477-10,10V482
                                c0,5.514-4.486,10-10,10h-364c-5.514,0-10-4.486-10-10V30c0-5.514,4.486-10,10-10h364c5.514,0,10,4.486,10,10v108.75
                                C404.007,144.273,408.485,148.75,414.007,148.75z"/>
                            </g>
                          </g>
                          <g>
                            <g>
                              <path d="M212.007,54c-50.729,0-92,41.271-92,92c0,26.317,11.11,50.085,28.882,66.869c0.333,0.356,0.687,0.693,1.074,1
                                c16.371,14.979,38.158,24.13,62.043,24.13c23.885,0,45.672-9.152,62.043-24.13c0.387-0.307,0.741-0.645,1.074-1
                                c17.774-16.784,28.884-40.552,28.884-66.869C304.007,95.271,262.736,54,212.007,54z M212.007,218
                                c-16.329,0-31.399-5.472-43.491-14.668c8.789-15.585,25.19-25.332,43.491-25.332c18.301,0,34.702,9.747,43.491,25.332
                                C243.405,212.528,228.336,218,212.007,218z M196.007,142v-6.5c0-8.822,7.178-16,16-16s16,7.178,16,16v6.5c0,8.822-7.178,16-16,16
                                S196.007,150.822,196.007,142z M269.947,188.683c-7.375-10.938-17.596-19.445-29.463-24.697c4.71-6.087,7.523-13.712,7.523-21.986
                                v-6.5c0-19.851-16.149-36-36-36s-36,16.149-36,36v6.5c0,8.274,2.813,15.899,7.523,21.986
                                c-11.867,5.252-22.088,13.759-29.463,24.697c-8.829-11.953-14.06-26.716-14.06-42.683c0-39.701,32.299-72,72-72s72,32.299,72,72
                                C284.007,161.967,278.776,176.73,269.947,188.683z"/>
                            </g>
                          </g>
                          <g>
                            <g>
                              <path d="M266.007,438h-54c-5.522,0-10,4.477-10,10s4.478,10,10,10h54c5.522,0,10-4.477,10-10S271.529,438,266.007,438z"/>
                            </g>
                          </g>
                          <g>
                            <g>
                              <path d="M266.007,382h-142c-5.522,0-10,4.477-10,10s4.478,10,10,10h142c5.522,0,10-4.477,10-10S271.529,382,266.007,382z"/>
                            </g>
                          </g>
                          <g>
                            <g>
                              <path d="M266.007,326h-142c-5.522,0-10,4.477-10,10s4.478,10,10,10h142c5.522,0,10-4.477,10-10S271.529,326,266.007,326z"/>
                            </g>
                          </g>
                          <g>
                            <g>
                              <path d="M88.366,272.93c-1.859-1.86-4.439-2.93-7.079-2.93c-2.631,0-5.211,1.07-7.07,2.93c-1.86,1.86-2.93,4.44-2.93,7.07
                                s1.069,5.21,2.93,7.07c1.87,1.86,4.439,2.93,7.07,2.93c2.64,0,5.21-1.07,7.079-2.93c1.86-1.86,2.931-4.44,2.931-7.07
                                S90.227,274.79,88.366,272.93z"/>
                            </g>
                          </g>
                          <g>
                            <g>
                              <path d="M88.366,328.93c-1.869-1.86-4.439-2.93-7.079-2.93c-2.631,0-5.2,1.07-7.07,2.93c-1.86,1.86-2.93,4.44-2.93,7.07
                                s1.069,5.21,2.93,7.07c1.87,1.86,4.439,2.93,7.07,2.93c2.64,0,5.21-1.07,7.079-2.93c1.86-1.86,2.931-4.44,2.931-7.07
                                S90.227,330.79,88.366,328.93z"/>
                            </g>
                          </g>
                          <g>
                            <g>
                              <path d="M88.366,384.93c-1.869-1.86-4.439-2.93-7.079-2.93c-2.631,0-5.2,1.07-7.07,2.93c-1.86,1.86-2.93,4.44-2.93,7.07
                                s1.069,5.21,2.93,7.07c1.859,1.86,4.439,2.93,7.07,2.93c2.64,0,5.22-1.07,7.079-2.93c1.86-1.86,2.931-4.44,2.931-7.07
                                S90.227,386.79,88.366,384.93z"/>
                            </g>
                          </g>
                          <g>
                            <g>
                              <path d="M266.007,270h-142c-5.522,0-10,4.477-10,10s4.478,10,10,10h142c5.522,0,10-4.477,10-10S271.529,270,266.007,270z"/>
                            </g>
                          </g>
                          <g>
                            <g>
                              <path d="M491.002,130.32c-9.715-5.609-21.033-7.099-31.871-4.196c-10.836,2.904-19.894,9.854-25.502,19.569L307.787,363.656
                                c-0.689,1.195-1.125,2.52-1.278,3.891l-8.858,79.344c-0.44,3.948,1.498,7.783,4.938,9.77c1.553,0.896,3.278,1.34,4.999,1.34
                                c2.092,0,4.176-0.655,5.931-1.948l64.284-47.344c1.111-0.818,2.041-1.857,2.73-3.052l125.841-217.963
                                C517.954,167.638,511.058,141.9,491.002,130.32z M320.063,426.394l4.626-41.432l28.942,16.71L320.063,426.394z M368.213,386.996
                                l-38.105-22l100.985-174.91l38.105,22L368.213,386.996z M489.054,177.693l-9.857,17.073l-38.105-22l9.857-17.073
                                c2.938-5.089,7.682-8.729,13.358-10.25c5.678-1.522,11.606-0.74,16.694,2.198c5.089,2.938,8.729,7.682,10.25,13.358
                                C492.772,166.675,491.992,172.604,489.054,177.693z"/>
                            </g>
                          </g>
                        </svg>
                      </span>
                    </a>
                  </td> 
                </tr>
              )}
            </table>
          </div>

          <div id="popup1" className={ styles.overlay}>
            <div className={styles.popup}>
              <h2>Détails :</h2>
              <a className={styles.close} href="#">&times;</a>
              <div className={styles.content}>
                <table>
                  <tbody>
                    <tr>
                      <td >Nom de l'employé :</td>
                      <td>{this.state.popUpUsername}</td>
                    </tr>
                    <tr>
                      <td >Motif d'absence :</td>
                      <td>{this.state.popUpMotif}</td>
                    </tr>
                    <tr>
                      <td >Date de début :</td>
                      <td>{this.state.popUpDA}</td>
                    </tr>
                    <tr>
                      <td >Date de fin :</td>
                      <td>{this.state.popUpDE}</td>
                    </tr>
                    <tr>
                      <td >Jours de congés :</td>
                      <td>{this.state.NumberOfDays}</td>
                    </tr>
                    <tr>
                      <td >Solde de départ :</td>
                      <td>{this.state.Solde}</td>
                    </tr>
                    <tr>
                      <td>Approbation du manager 1 :</td>
                      <td>{this.state.popUpApprobation}</td>
                    </tr>
                    <tr>
                      <td>Statut d'approbation :</td>
                      <td>{this.state.popUpStatus}</td>
                    </tr>
                    <tr>
                      <td>Commentaire :</td>
                      <td>{this.state.popUpComment}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>


          <div className={styles.paginations}>
            <button className={styles.pagination}>Prev</button>
            <span id="page"></span>
            <button className={styles.pagination}>Next</button>
          </div>
          
          
        </div>
    );
  }
}
