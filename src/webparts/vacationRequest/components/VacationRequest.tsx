import * as React from 'react';
import { addDays } from 'date-fns';
import styles from './VacationRequest.module.scss';
import stylescustom from './VacationRequestCustom.module.scss';
import { IVacationRequestProps } from './IVacationRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Dropdown, IDropdownOption, IDropdownProps, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { DatePicker, IDatePickerStrings, IDatePickerStyles, IDatePickerStyleProps } from 'office-ui-fabric-react/lib/DatePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { sp, Web, IWeb } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/attachments";
import { getTheme } from "@uifabric/styling";
// var img = require('../../../image/UCT_image.png');
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import SweetAlert from 'sweetalert2-react';
import {
  Fabric,
  DefaultButton,
  Slider,
  Panel,
  PanelType,
  loadTheme
} from "office-ui-fabric-react";
import { MSGraphClient, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import Holidays from 'date-holidays'
export interface IUser {
  displayName: string;
  mail:string;
}
import languages from "../../../languages/languages.json"


//import { IUserPresenceState } from './IVacationRequestProps';


// import { IPresence } from "../../../model/IPresence";
// import GraphService from '../../../services/GraphService';
// import { getTheme } from "@uifabric/styling";
// import { JSONParser } from '@pnp/odata';







export default class VacationRequest extends React.Component<IVacationRequestProps, {}> {

  // state of vacation request webpart
  public state = {

    ////////// data of current user ///////////////////
    currentUserDisplayName: '',
    currentUserID: '',
    currentUserMail: '',
    currentUserPrincipalName: '',
    //////////////////////////////////////////////////




    ///////////////////////////////  vacation data  /////////////////////////
    motifAbsence : "",  // option of absence
    decesOptionData: "", // other options when user select 'Décès' in Motif d'absence
    mariageOtionData: "", // other options when user select 'Mariage' in Motif d'absence
    DateDebut: new Date(), // begin date of Vacation
    DateFin: new Date(), // end date of vacation
    fileName: "", // file name upload
    file: {}, // file upload information
    comment: "", // comment of form
    replacedBy: [] , // data of user who replaces you on vacation
    numberOfVacationDays: 0,  // number of days in vacation
    numberOfVacationDaysForUser: 0,
    ////////////////////////////////////////////////////////////////////////



    ////////// Enabled and disabled buttons and alerts for dynamic form ///////////////
    disabledDays: false,
    disableSubmitButton: true,
    alertShowed: false,
    ///////////////////////////////////////////////////////////////////////////////////

    errorMessage : "",


    ///////////////////////////// language configuration //////////////////////////////
    LanguageSelected: 0,
    ///////////////////////////////////////////////////////////////////////////////////

    ///////////////////////////// file Language Data   ////////////////////////////////
    Congepaye: "",
    DemiJournee: "",
    Maladie: "",
    Naissance: "",
    Mariage: "",
    Deces: "",
    Circonsion: "",
    Parents: "",
    Conjoint: "",
    Enfants: "",
    GrandParent: "",
    Freres: "",
    Soeurs:"",
    PetitEnfants: "",
    MariageEnfant: "",

    TitreDuPage: "",
    EmployeeName: "",
    EmailOrganisation: "",
    EmployeeID: "",
    Email: "",
    Champs: "",
    Reason1: "",
    Reason2: "",
    StartDate: "",
    EndDate: "",
    Attach: "",
    Jours: "",
    RemplacePar: "",
    CommentFile: "",
    OtherDetails: "",
    Solde: "",
    Enregistrer: ""
    ///////////////////////////////////////////////////////////////////////////////////


  };
  public user: IUser[] = [];






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
          this.setState({
            currentUserDisplayName: userData[0].displayName,
            currentUserID: userData[0].id,
            currentUserMail: userData[0].mail,
            currentUserPrincipalName: userData[0].userPrincipalName
          }) 
        })
    })
  }
  


  // Condition for disable endDate if user select "half day","birth","Mariage","Décès" or "Circoncision"
  private disableEndDate = () => {
    if ((this.state.motifAbsence === "half day") || (this.state.motifAbsence === 'birth') || (this.state.motifAbsence === 'wedding') || ((this.state.motifAbsence === 'death')) || ((this.state.motifAbsence === 'Circumcision'))) {
      return true;
    }
    return false;
  }



  // enable and disable submit button
  private disabledSubmitButton = () => {
    var test = true
    switch (this.state.motifAbsence) {
      // 1
      case "half day":
      case "birth":
      case "Circumcision":
        // test DA, remplacer par
        if ((this.state.DateDebut !== null)&&(this.state.replacedBy.length > 0)){
          test = false
        }
      break;

      // 2
      case "wedding":
      case "death":
        // test Décès options
        // test DA, REMP, Plus details
        if (this.state.motifAbsence === "death"){
          if ((this.state.DateDebut !== null) && (this.state.replacedBy.length > 0) && (this.state.decesOptionData !== "")){
            test = false
          }
        }
        // test Mariage options
        // test DA, REMP, Plus details
        if (this.state.motifAbsence === "wedding"){
          if ((this.state.DateDebut !== null) && (this.state.replacedBy.length > 0) && (this.state.mariageOtionData !== "")){
            test = false
          }

        }
      break;
      
      // 3
      case "illness":
        // test DA, DF, REMP, File
        if ((this.state.DateDebut !== null) && (this.state.DateFin !== null) && (this.state.replacedBy.length > 0) && (this.state.fileName !== "") && (this.state.numberOfVacationDays > 0)){
          test = false
        }
      break;
      
      // 4
      case "paid vacation":
        // tester DA, DF, REMP 
        if ((this.state.DateDebut !== null) && (this.state.DateFin !== null) && (this.state.replacedBy.length > 0) && (this.state.numberOfVacationDays > 0)){
          test = false
        }
      break; 
    };

    return test
  }



  // initialise disable end date for default vacation types
  public ChangeDefaultEndDateDisabled = (beginDate) => {
    if (this.state.disabledDays){
      var defaultNumberOfVacationDay, endDate
      switch (this.state.motifAbsence) {
        case 'half day':
          defaultNumberOfVacationDay = 0.5;
          endDate = beginDate
          break;

        case 'birth':
          defaultNumberOfVacationDay = 2;
          endDate = addDays(beginDate,defaultNumberOfVacationDay)
        break;

        case 'Circumcision':
          defaultNumberOfVacationDay = 1;
          endDate = addDays(beginDate,defaultNumberOfVacationDay)
        break;

        case 'wedding':
          if (this.state.mariageOtionData !== ""){
            defaultNumberOfVacationDay = this.state.numberOfVacationDays;
            endDate = addDays(beginDate,defaultNumberOfVacationDay)
          }
          
        break;

        case 'death':
          if (this.state.decesOptionData !== ""){
            defaultNumberOfVacationDay = this.state.numberOfVacationDays;
            endDate = addDays(beginDate,defaultNumberOfVacationDay)
          }

        break;
      }
      this.setState({DateFin:endDate})
    }
  }


  
  // setstate of begin and end date
  private SelectDate = (Day,Month,Year,beginDate) => {
    
    // if beginDate param equal to true -> this function setstate the beginDate 
    if (beginDate){
      // convert the type of date seleted
      const date = Year.toString() + "-" + Month.toString() + "-" + Day.toString()  + " GMT";
      const newDateFormat = new Date(date);
      if (this.state.motifAbsence === "paid vacation" || this.state.motifAbsence === "illness") {
        this.setState({DateDebut:newDateFormat, DateFin:null, numberOfVacationDays:0});
      }else {
        this.setState({DateDebut:newDateFormat});
      }
      
      this.ChangeDefaultEndDateDisabled(newDateFormat)

    // if beginDate param equal to false -> this function setstate the endDate
    }else {
      // convert the type of date
      const date = Year.toString() + "-" + Month.toString() + "-" + Day.toString()  + " GMT";
      
      const newDateFormat = new Date(date);
      // const diffDays = this.SumVacationDays(this.state.DateDebut, newDateFormat)
      const numberOfDays = this.getNumberOfDays(this.state.DateDebut, newDateFormat)

      this.setState({DateFin:newDateFormat, numberOfVacationDays:numberOfDays});
    }
      
  }


  // public SumVacationDays = (dateDebut, DateFin) => {
  //   var diffDays = DateFin.getTime() - dateDebut.getTime();
  //   diffDays = diffDays / (1000 * 3600 * 24);
  //   return diffDays
  // }



  



  // upload file to vaquation request form
  public addFile(fileInfo) {
    this.setState({ fileName: fileInfo.target.files[0].name });
    this.setState({ file: fileInfo.target.files[0] });
  }

  


  // initialise file input if user delete files selected
  public initImage() {
    this.setState({ file: {}, fileName:"" });
    (document.getElementById('uploadFile') as HTMLInputElement).value = "";
  }



  // setstate changes of comment input
  public handleChange = (event) => {
    this.setState({comment:event.target.value});
  }



  // get new user when we change a new user in "Remplacé Par" 
  public _getPeoplePickerItems = async (items: any[]) => {
    if (items.length > 0) {
      if (items[0].id && items[0].text && items[0].secondaryText){
        let replacedUserData = {ID:items[0].id,name:items[0].text,email:items[0].secondaryText}
        this.setState({replacedBy:[replacedUserData]})
      }
    }else {
      this.setState({replacedBy:[]})
    }
  }


  // calculate number of days in vacation 
  public vacationDays = () => {
    var diffDays = this.state.DateFin.getTime() - this.state.DateDebut.getTime()
    diffDays = diffDays / (1000 * 3600 * 24);
    this.setState({numberOfVacationDays:diffDays})
  }






  // Setstate the default of vacation days if the user select "Décès"
  public defaultVacationDays = () => {
    var defaultNumberOfVacationDay = 0
    var disabledDays = false
    if (this.state.motifAbsence !== ""){
      switch (this.state.motifAbsence) {
        case 'half day':
          defaultNumberOfVacationDay = 0.5;
          disabledDays = true
          break;
        case 'birth':
          defaultNumberOfVacationDay = 2;
          disabledDays = true
        break;
        case 'Circumcision':
          defaultNumberOfVacationDay = 1;
          disabledDays = true
        break;
      }
    }
    this.setState({DateDebut:null, DateFin:null, disabledDays:disabledDays, numberOfVacationDays:defaultNumberOfVacationDay})
  }


  // Setstate the default of vacation days if the user select "Décès"
  public defaultVacationDaysMariage = () => {
    var defaultNumberOfVacationDay = 0
    if (this.state.mariageOtionData !== ""){
      switch (this.state.mariageOtionData) {
        case 'wedding':
          defaultNumberOfVacationDay = 3;
          break;
        case 'child marriage':
          defaultNumberOfVacationDay = 1;
        break;
      }

      this.setState({disabledDays:true, DateDebut:null, DateFin:null })
    }
    this.setState({numberOfVacationDays:defaultNumberOfVacationDay})

  }




  // Setstate the default of vacation days if the user select "Mariage"
  public defaultVacationDaysDeces = () => {
    var defaultNumberOfVacationDay = 0
    if (this.state.decesOptionData !== ""){
      switch (this.state.decesOptionData) {
        case 'Parents':
          defaultNumberOfVacationDay = 3;
          break;
        case 'Spouse':
          defaultNumberOfVacationDay = 3;
        break;
        case 'Childrens':
          defaultNumberOfVacationDay = 3;
        break;

        case 'Grand-parents':
          defaultNumberOfVacationDay = 2;
        break;
        case 'Brothers':
          defaultNumberOfVacationDay = 2;
        break;
        case 'sisters':
          defaultNumberOfVacationDay = 2;
        break;
        case 'small-child':
          defaultNumberOfVacationDay = 2;
        break; 
      }
      this.setState({disabledDays:true, DateDebut:null, DateFin:null})
    }
    this.setState({numberOfVacationDays:defaultNumberOfVacationDay})
  }



  // send All form data to a sharepoint list
  public collectAllData = async () => {

    // Initialise data
    var formData
    var detailAbsence = ""
    if(this.state.decesOptionData !== "") detailAbsence = this.state.decesOptionData
    if(this.state.mariageOtionData !== "") detailAbsence = this.state.mariageOtionData

    if (this.state.motifAbsence === "paid vacation" || this.state.motifAbsence === "half day"  || this.state.motifAbsence === "birth" || this.state.motifAbsence === "Circumcision" ){
      formData = {
        'Comment': this.state.comment,
        'EndDate': this.state.DateFin,
        'DetailMotifAbsence': "",
        'NrbDays': this.state.numberOfVacationDays,
        'RemainingDays': "0",
        'ReplacedById': this.state.replacedBy[0].ID,
        'ReplacedByStringId': this.state.replacedBy[0].ID.toString(),
        'RequestType': "en cours",
        'ctgVacation': this.state.motifAbsence,
        'dateDeDepart': this.state.DateDebut,
        'vacationType': "",
      };
    }else {
      formData = {
        'Comment': this.state.comment,
        'EndDate': this.state.DateFin,
        'DetailMotifAbsence': this.state.motifAbsence +" "+detailAbsence,
        'NrbDays': this.state.numberOfVacationDays,
        'RemainingDays': "0",
        'ReplacedById': this.state.replacedBy[0].ID,
        'ReplacedByStringId': this.state.replacedBy[0].ID.toString(),
        'RequestType': "en cours",
        'ctgVacation': this.state.motifAbsence,
        'dateDeDepart': this.state.DateDebut,
        'vacationType': detailAbsence,
      };
    }

    
    // add new vacation request to sharepoint list 
    const sendData = await Web(this.props.url).lists.getByTitle('vacationRequest').items.add(formData);

    // send attachement file in the new item if user add a new file
    if (this.state.fileName !== ''){
      const item = Web(this.props.url).lists.getByTitle('vacationRequest').items.getById(sendData.data.ID);
      const result = await item.attachmentFiles.add(this.state.fileName,"add file");
      // console.log(result)
    }

    this.setState({alertShowed:true})
  }



  // Get numbers of vacation days for current User
  public numberOfVacationDays = async() => {
    if (this.state.currentUserID !== ""){
      const items = await Web(this.props.url).lists.getByTitle('usersVacationDays').items();
      var currentUserVacationDays = items.filter(item => this.state.currentUserID === item.ID_user);
      this.setState({numberOfVacationDaysForUser:currentUserVacationDays[0].number_of_vacation})
    }

  }


  // Get number of days between two dates exclude weekends
  public getNumberOfDays = (startDate, endDate) => {
    var iWeeks, iDateDiff, iAdjust = 0;
    var iWeekday1 = startDate.getDay();
    var iWeekday2 = endDate.getDay();
    iWeekday1 = (iWeekday1 == 0) ? 7 : iWeekday1; // change Sunday from 0 to 7
    iWeekday2 = (iWeekday2 == 0) ? 7 : iWeekday2;
    if ((iWeekday1 > 5) && (iWeekday2 > 5)) iAdjust = 1; // adjustment if both days on weekend
    iWeekday1 = (iWeekday1 > 5) ? 5 : iWeekday1; // only count weekdays
    iWeekday2 = (iWeekday2 > 5) ? 5 : iWeekday2;
 
    // calculate differnece in weeks (1000mS * 60sec * 60min * 24hrs * 7 days = 604800000)
    iWeeks = Math.floor((endDate.getTime() - startDate.getTime()) / 604800000)
 
    if (iWeekday1 <= iWeekday2) {
      iDateDiff = (iWeeks * 5) + (iWeekday2 - iWeekday1)
    } else {
      iDateDiff = ((iWeeks + 1) * 5) - (iWeekday1 - iWeekday2)
    }
 
    iDateDiff -= iAdjust // take into account both days on weekend
    return iDateDiff
  }
  


  // get all hoidays of current year 
  public getHolidayDays = () => {
    var hd = new Holidays('TN');
    var currentYear = new Date();
    var currentYearHolidays = hd.getHolidays(currentYear.getFullYear());
    // console.log(currentYearHolidays)
  }

  public initialiseCurrentLanguage = () => {
    var languageSelectedID = this.props.LanguageSelected
    // console.log(languages["frenshOptions"]['TitreDuPage'])

    switch (languageSelectedID) {     ////// A completer 
      // Frensh language
      case 1:
        // update file text
        this.setState({
          Congepaye: languages["frenshOptions"]['Congé payé'],
          DemiJournee: languages["frenshOptions"]['Demi journée'],
          Maladie: languages["frenshOptions"]['Maladie'],
          Naissance: languages["frenshOptions"]['Naissance'],
          Mariage: languages["frenshOptions"]['Mariage'],
          Deces: languages["frenshOptions"]['Décès'],
          Circonsion: languages["frenshOptions"]['Congé'],
          Parents: languages["frenshOptions"]['Congé'],
          Conjoint: languages["frenshOptions"]['Congé'],
          Enfants: languages["frenshOptions"]['Congé'],
          GrandParent: languages["frenshOptions"]['Congé'],
          Freres: languages["frenshOptions"]['Congé'],
          Soeurs: languages["frenshOptions"]['Congé'],
          PetitEnfants: languages["frenshOptions"]['Congé'],
          MariageEnfant: languages["frenshOptions"]['Congé'],

          TitreDuPage: languages["frenshOptions"]['Congé'],
          EmployeeName: languages["frenshOptions"]['Congé'],
          EmailOrganisation: languages["frenshOptions"]['Congé'],
          EmployeeID: languages["frenshOptions"]['Congé'],
          Email: languages["frenshOptions"]['Congé'],
          Champs: languages["frenshOptions"]['Congé'],
          Reason1: languages["frenshOptions"]['Congé'],
          Reason2: languages["frenshOptions"]['Congé'],
          StartDate: languages["frenshOptions"]['Congé'],
          EndDate: languages["frenshOptions"]['Congé'],
          Attach: languages["frenshOptions"]['Congé'],
          Jours: languages["frenshOptions"]['Congé'],
          RemplacePar: languages["frenshOptions"]['Congé'],
          CommentFile: languages["frenshOptions"]['Congé'],
          OtherDetails: languages["frenshOptions"]['Congé'],
          Solde: languages["frenshOptions"]['Congé'],
          Enregistrer: languages["frenshOptions"]['Congé']
        })
      break;

      // Arabic Language
      case 2:
        console.log(2)

      break;
      
      // English Language
      case 3:
        console.log(3)

      break;
    };

  }



  // update Get Users Stat when initialise page
  componentDidMount(): void {
    this.getUsers();
    this.initialiseCurrentLanguage();
  }





  public render(): React.ReactElement<IVacationRequestProps> {
    
    // props of webpart
    const {
      description,
      url,
      context
    } = this.props;
    
    // Style of Dropdown
    const dropdownStyles: Partial<IDropdownStyles> = {
      title: { backgroundColor: "white" },
    };

    // style of inputs
    const controlClass = mergeStyleSets({
      TextField: { backgroundColor: "white", }
    });


    // // options of Absence
    // const motifAbsence = [
    //   {key: "paid vacation",text: "paid vacation"},
    //   {key: "half day", text: "half day"}, 
    //   {key: "illness", text: "illness"}, 
    //   {key: "birth", text: "birth"}, 
    //   {key: "Mariage", text: "Mariage"}, 
    //   {key: "Décès", text: "Décès"}, 
    //   {key: "Circoncision", text: "Circoncision"},
    // ];
    
    // options of Absence
    const motifAbsence = [
      {key: "paid vacation",text: "paid vacation"},
      {key: "half day", text: "half day"}, 
      {key: "illness", text: "illness"}, 
      {key: "birth", text: "birth"}, 
      {key: "wedding", text: "wedding"}, 
      {key: "death", text: "death"}, 
      {key: "Circumcision", text: "Circumcision"},
    ];


    // // the other options when user choice is "Décès"
    // const decesOptions = [
    //   { key: "Parents", text: "Parents", },
    //   { key: "Conjoint", text: "Conjoint", },
    //   { key: "Enfants", text: "Enfants", },
    //   { key: "Grands-parents", text: "Grands-parents", },
    //   { key: "Frères", text: "Frères", },
    //   { key: "Sœurs", text: "Sœurs", },
    //   { key: "Petits-enfants", text: "Petits-enfants", }
    // ];

    // the other options when user choice is "Décès"
    const decesOptions = [
      { key: "Parents", text: "Parents", },
      { key: "Spouse", text: "Spouse", },
      { key: "Childrens", text: "Childrens", },
      { key: "Grand-parents", text: "Grand-parents", },
      { key: "Brothers", text: "Brothers", },
      { key: "sisters", text: "sisters", },
      { key: "small-child", text: "small-child", }
    ];


    // // the other options when user choice is "Mariage"
    // const mariageOptions =  [
    //   { key: "Mariage", text: "Mariage", },
    //   { key: "Mariage d’un enfant", text: "Mariage d’un enfant", }
    // ];

    // the other options when user choice is "Mariage"
    const mariageOptions =  [
      { key: "wedding", text: "wedding", },
      { key: "child marriage", text: "child marriage", }
    ];


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

    // get the current date 
    const CurrentDate = new Date().getDate().toString() + "/" + (new Date().getMonth()+1).toString() + "/" + new Date().getFullYear().toString();
    

    // get the theme of sharepoint
    const theme = getTheme();

    this.numberOfVacationDays();   // Get number of vacation days for current user
    // this.getHolidayDays();     // get holidays days of year in current country
    

    return (
      <Fabric
        className="App"
        style={{ background: theme.semanticColors.bodyBackground, color: theme.semanticColors.bodyText }}
      >
        <div className={stylescustom.vacationRequest}>
          <div className={stylescustom.DC}>
            <p className={stylescustom.datenow}>Date : <span className="date-time">{CurrentDate}</span></p>
            {/* <div className={stylescustom.titleh1}>Demande de congé </div> */}
            <div className={stylescustom.titleh1}>Leave request {this.props.LanguageSelected}</div>
            <div className={stylescustom.line}></div>


            <div className={stylescustom.row}>
              <div className={stylescustom.col}>
                <table className={stylescustom.table}>
                  <tbody>
                    <tr>
                      {/* <td className={stylescustom.key}>Nom de l'employé</td> */}
                      <td className={stylescustom.key}>employee name</td>
                      <td className={stylescustom.value}>{this.state.currentUserDisplayName} </td>
                    </tr>
                    <tr>
                      {/* <td className={stylescustom.key}>Adresse email de l'organisation</td> */}
                      <td className={stylescustom.key}>Organization email address</td>
                      <td className={stylescustom.value}>{this.state.currentUserPrincipalName}</td>
                    </tr>
                    <tr>
                      {/* <td className={stylescustom.key}>ID employé</td> */}
                      <td className={stylescustom.key}>Employee ID</td>
                      <td className={stylescustom.value}>{this.state.currentUserID}</td>
                    </tr>
                    <tr>
                      {/* <td className={stylescustom.key}>Adresse email</td> */}
                      <td className={stylescustom.key}>E-mail address</td>
                      <td className={stylescustom.value}>{this.state.currentUserMail}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>


            {/* <p className={stylescustom.indique}>* Indique un champ obligatoire</p> */}
            <p className={stylescustom.indique}>* Required field</p>
            <div className={stylescustom.row}>

              {/* Select absence Motif */}
              <div className={stylescustom.data}>
                {/* <p className={stylescustom.title}>* Motif d'absence :</p> */}
                <p className={stylescustom.title}>* Reason for absence :</p>
                <Dropdown
                  styles={dropdownStyles}
                  // onRenderTitle={this.onRenderTitle}
                  // onRenderOption={this.onRenderOption}
                  // onRenderCaretDown={this.onRenderCaretDown}
                  options={motifAbsence}
                  onChanged={(value) => this.setState({ motifAbsence:value['key'], disabledDays:false},this.defaultVacationDays)}
                  defaultSelectedKey={this.state.motifAbsence}
                />
              </div>



              {/* ********* Show other décès options when user select décès in motif d'absence ********* */}
              {this.state.motifAbsence == 'death' && <div className={stylescustom.data}>
                {/* <p className={stylescustom.title}>* Plus de détails</p> */}
                <p className={stylescustom.title}>* More details</p>
                <Dropdown
                  styles={dropdownStyles}
                  // onChange={this.onSelectionChanged}
                  // onRenderTitle={this.onRenderTitle}
                  // onRenderOption={this.onRenderOption}
                  // onRenderCaretDown={this.onRenderCaretDown}
                  options={decesOptions}
                  onChanged={(value) => this.setState({decesOptionData:value['key']}, this.defaultVacationDaysDeces)}
                  defaultSelectedKey={this.state.decesOptionData}
                  // errorMessage={this.state.errors.VacationType !== "" ? this.state.errors.VacationType : ""}
                />
              </div>}
              {/* ******************************************************************************************* */}




              {/* ********* Show other mariage options when user select mariage in motif d'absence ********* */}
              {this.state.motifAbsence == 'wedding' && <div className={stylescustom.data}>
                {/* <p className={stylescustom.title}>* Plus de détails</p> */}
                <p className={stylescustom.title}>* More details</p>
                <Dropdown
                  styles={dropdownStyles}
                  // onChange={this.onSelectionChanged}
                  // onRenderTitle={this.onRenderTitle}
                  // onRenderOption={this.onRenderOption}
                  // onRenderCaretDown={this.onRenderCaretDown}
                  options={mariageOptions}
                  onChanged={(value) => this.setState({mariageOtionData:value['key']},this.defaultVacationDaysMariage)}
                  defaultSelectedKey={this.state.mariageOtionData}
                  // errorMessage={this.state.errors.VacationType !== "" ? this.state.errors.VacationType : ""}
                />
              </div>}
              {/* ******************************************************************************************* */}


              <div className={stylescustom.data}>
                {/* <p className={stylescustom.title}>* Date debut :</p> */}
                <p className={stylescustom.title}>* Start date :</p>
                <DatePicker
                  className={controlClass.TextField}
                  allowTextInput={false}
                  strings={DatePickerStrings}
                  value={this.state.DateDebut}
                  onSelectDate={(e) => { this.SelectDate(e.getDate(), e.getMonth()+1 ,e.getFullYear(), true) }}
                  ariaLabel="Select a date"
                  minDate={new Date()} 
                //this.dateDiffInDays(this.state.StartDate,this.state.EndDate)
                />
              </div>



              <div className={stylescustom.data}>
                {/* <p className={stylescustom.title}>* Date Fin :</p> */}
                <p className={stylescustom.title}>* End date :</p>
                <DatePicker
                  className={controlClass.TextField}
                  allowTextInput={false}
                  strings={DatePickerStrings}
                  value={this.state.DateFin}
                  onSelectDate={(e) => { this.SelectDate(e.getDate(), e.getMonth()+1 ,e.getFullYear(), false) }}
                  ariaLabel="Select a date"
                  minDate={this.state.DateDebut}
                  disabled={this.disableEndDate()}
                />
              </div>
              


              <div className={stylescustom.data}>
                <p className={stylescustom.title}>
                  {this.state.motifAbsence === "illness" && <span>*</span>}Attach a supporting document :
                </p>
                {/* <label htmlFor="uploadFile" className={stylescustom.btn}>Choisir un élément</label> */}
                <label htmlFor="uploadFile" className={stylescustom.btn}>Choose an item</label>
                <input type="file" id="uploadFile" style={{ display: 'none' }}
                  accept=".jpg, .jpeg, .png , .pdf , .doc ,.docx"
                  onChange={(e) => { this.addFile(e); }} 
                />
                {this.state.file && <span style={{ marginLeft: 10, fontSize: 14 }}>{this.state.fileName} <span style={{ cursor: 'pointer' }} onClick={() => { this.initImage(); }}>&#10006;</span></span>}
                {/* <span style={{ color: "rgb(168, 0, 0)", fontSize: 12, fontWeight: 400, display: 'block' }}>
                  {this.state.errors.file !== "" ? this.state.errors.file : ""}
                </span> */}
              </div>

            </div>



            <div className={stylescustom.row}>
              <div className={stylescustom.data}>
                {/* <p className={stylescustom.title}>Jours :</p> */}
                <p className={stylescustom.title}>Days :</p>

                {/* If user Select a vacation d with default days */}
                {this.state.disabledDays && <TextField className={controlClass.TextField} disabled={this.state.disabledDays} value={this.state.numberOfVacationDays.toString()} />}

                {/* if user select a vacation with choice days */}
                {!this.state.disabledDays && this.state.numberOfVacationDays}


                {/* <TextField className={controlClass.TextField} disabled={true} value={this.state.numberOfDayVacationQty} />
                <span style={{ color: "rgb(168, 0, 0)", fontSize: 12, fontWeight: 400, display: 'block' }}>
                  {this.state.errors.numberOfDayVacationQty !== "" ? this.state.errors.numberOfDayVacationQty : ""}
                </span> */}
                {/* {!this.disableEndDate() && <TextField style={{color: "rgb(168, 0, 0)"}} className={controlClass.TextField} disabled={true} value={this.state.numberOfVacationDays.toString()} />}
                {this.disableEndDate() && <TextField className={controlClass.TextField} disabled={true} value={this.state.numberOfVacationDays.toString()} />} */}

              </div>

              
            </div>
            



            <div className={stylescustom.row}>
              <div className={stylescustom.datarem}>
                {/* <p className={stylescustom.title}>Remplacé par :</p> */}
                <p className={stylescustom.title}>Replaced by :</p>
                <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={1}
                  required={false}
                  onChange={this._getPeoplePickerItems}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  ensureUser={true}
                />
                <svg className={stylescustom.iconsearch} xmlns="http://www.w3.org/2000/svg" width="20" height="20.003" viewBox="0 0 20 20.003">
                  <path data-name="Icon awesome-search" d="M19.728,17.294,15.833,13.4a.937.937,0,0,0-.664-.273h-.637a8.122,8.122,0,1,0-1.406,1.406v.637a.937.937,0,0,0,.273.664l3.895,3.895a.934.934,0,0,0,1.324,0l1.106-1.106a.942.942,0,0,0,0-1.328Zm-11.6-4.168a5,5,0,1,1,5-5A5,5,0,0,1,8.126,13.126Z" />
                </svg>
              </div>
            </div>



            <div className={stylescustom.row}>
              <div className={stylescustom.comment}>
                {/* <p className={stylescustom.title}>Commentaire :</p> */}
                <p className={stylescustom.title}>Comment :</p>
                <TextField className={controlClass.TextField} value={this.state.comment} multiline onChange={this.handleChange} />
              </div>
            </div>


            <table className={stylescustom.ad}>
              <thead>
                {/* <th className={stylescustom.title} >Autres détails</th> */}
                <th className={stylescustom.title} >Other details</th>
              </thead>
              <tbody className={stylescustom.tbody}>
                <tr>
                  {/* <td className={stylescustom.key}>Solde des congés </td> */}
                  <td className={stylescustom.key}>Leave balance </td>
                  <td className={stylescustom.value}>{this.state.numberOfVacationDaysForUser}</td>
                </tr>
              </tbody>
            </table>



            <div className={stylescustom.btncont}>
              {/* {this.state.loadingFile ? <Spinner size={SpinnerSize.large} className={stylescustom.spinner} /> : ""} */}
              {/* <button className={stylescustom.btn} onClick={()=>this.collectAllData()} disabled={this.disabledSubmitButton()} >soumettre la demande</button> */}
              <button className={stylescustom.btn} onClick={()=>this.collectAllData()} disabled={this.disabledSubmitButton()} >submit request</button>
            </div>


            {/* <SweetAlert
            show={this.state.alertShowed} title="Demande de congé" text="Demande envoyée"
            confirmButtonColor='#7D2935'
            onConfirm={() => window.open(this.props.url + "/SitePages/Vacation-List.aspx", "_self")}
            imageWidth="200"
            imageHeight="200"
            /> */}
            <SweetAlert
            show={this.state.alertShowed} title="Leave request" text="Request submited"
            confirmButtonColor='#7D2935'
            onConfirm={() => window.open(this.props.url + "/SitePages/Vacation-List.aspx", "_self")}
            imageWidth="200"
            imageHeight="200"
            />



            {/* <SweetAlert
              show={this.state.alert} title="Demande de congé" text="Demande envoyée"
              imageUrl={img}
              confirmButtonColor='#7D2935'
              onConfirm={() => window.open(this.props.webURL + "/SitePages/Tableau-de-bord-utilisateur-des-demandes-de-congé.aspx", "_self")}
              imageWidth="200"
              imageHeight="200"
            />
            <SweetAlert
              show={this.state.alerteligibility} title="Demande de congé" text="Votre solde de congé est insuffisant"
              imageUrl={img}
              confirmButtonColor='#7D2935'
              onConfirm={() => this.setState({ alerteligibility: false })}
              // onConfirm={() => window.open(this.props.webURL + "/SitePages/Demande-de-congé.aspx", "_self")}
              imageWidth="200"
              imageHeight="200"
            /> */}
          </div>
        </div>
      </Fabric>
    );
  }
}
