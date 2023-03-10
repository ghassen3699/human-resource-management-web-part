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
    decesOptionData: "", // other options when user select 'D??c??s' in Motif d'absence
    mariageOtionData: "", // other options when user select 'Mariage' in Motif d'absence
    DateDebut: new Date(), // begin date of Vacation
    DateFin: new Date(), // end date of vacation
    fileName: "", // file name upload
    file: {}, // file upload information
    comment: "", // comment of form
    replacedBy: [] , // data of user who replaces you on vacation
    numberOfVacationDays: 0,  // number of days in vacation
    // numberOfVacationDaysForUser: 0,
    vacationDaysOfCurrentUser: 0,
    itemID: 0,
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
    Enregistrer: "",
    DateFile: "",
    directionFile: "",
    ChoisirFichier: "",
    TitreMessageValidation:"",
    TextMessageValidation: ""
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
          res.value.map(user => this.addUsersToListe(user))
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


  // Add users to users liste with your information(ID, displayName, mail and userDisplayName)
  private addUsersToListe = async (userInformation) => {
    const items = await sp.web.lists.getByTitle("usersVacationDays").items();
    const result = items.filter(user => user.UserID === userInformation.id)

    // Add new item to usersVacationDays sharepoint list if this algo ditect a new user in Azure AD
    if (result.length === 0){
      const iar = await sp.web.lists.getByTitle("usersVacationDays").items.add({
        UserID: userInformation.id,   // ID of item 
        DisplayName: userInformation.displayName,    // display name of item
        UserPrincipalName: userInformation.userPrincipalName,   // user principal name of item 
        UserMail: userInformation.mail,   // email of item
        v: userInformation.displayName,
        NumberOfDays:22
      });
    }
  }


  // Get numbers of vacation days for current User
  public getNumberOfVacationDays = async() => {
    if (this.state.currentUserID !== ""){
      const items = await Web(this.props.url).lists.getByTitle('usersVacationDays').items();
      // var currentUserVacationDays = items.filter(item => this.state.currentUserID === item.UserID);
      var currentUserVacationDays = items.filter(item => this.state.currentUserID === item.UserID);
      this.setState({vacationDaysOfCurrentUser: currentUserVacationDays[0].NumberOfDays, itemID: currentUserVacationDays[0].ID})
    }
  }


  // Condition for disable endDate if user select "half day","birth","Mariage","D??c??s" or "Circoncision"
  private disableEndDate = () => {
    if ((this.state.motifAbsence === this.state.DemiJournee) || (this.state.motifAbsence === this.state.Naissance) || (this.state.motifAbsence === this.state.Mariage) || ((this.state.motifAbsence === this.state.Deces)) || ((this.state.motifAbsence === this.state.Circonsion))) {
      return true;
    }
    return false;
  }



  // enable and disable submit button
  private disabledSubmitButton = () => {
    var test = true
    switch (this.state.motifAbsence) {
      // 1
      case this.state.DemiJournee:
      case this.state.Naissance:
      case this.state.Circonsion:
        // test DA, remplacer par
        // if ((this.state.DateDebut !== null)&&(this.state.replacedBy.length > 0)){
        //   test = false
        // }
        if ((this.state.DateDebut !== null)){
          test = false
        }
      break;

      // 2
      case this.state.Mariage:
      case this.state.Deces:
        // test D??c??s options
        // test DA, REMP, Plus details
        if (this.state.motifAbsence === this.state.Deces){
          // if ((this.state.DateDebut !== null) && (this.state.replacedBy.length > 0) && (this.state.decesOptionData !== "")){
          //   test = false
          // }
          if ((this.state.DateDebut !== null) && (this.state.decesOptionData !== "")){
            test = false
          }
        }
        // test Mariage options
        // test DA, REMP, Plus details
        if (this.state.motifAbsence === this.state.Mariage){
          // if ((this.state.DateDebut !== null) && (this.state.replacedBy.length > 0) && (this.state.mariageOtionData !== "")){
          //   test = false
          // }
          if ((this.state.DateDebut !== null) && (this.state.mariageOtionData !== "")){
            test = false
          }

        }
      break;
      
      // 3
      case this.state.Maladie:
        // test DA, DF, REMP, File
        // if ((this.state.DateDebut !== null) && (this.state.DateFin !== null) && (this.state.replacedBy.length > 0) && (this.state.fileName !== "") && (this.state.numberOfVacationDays > 0) && (this.calculerVacationDaysForSubmit(this.state.numberOfVacationDays,this.state.vacationDaysOfCurrentUser))){
        //   test = false
        // }
        if ((this.state.DateDebut !== null) && (this.state.DateFin !== null) && (this.state.fileName !== "") && (this.state.numberOfVacationDays > 0) && (this.calculerVacationDaysForSubmit(this.state.numberOfVacationDays,this.state.vacationDaysOfCurrentUser))){
          test = false
        }
      break;
      
      // 4
      case this.state.Congepaye:
        // tester DA, DF, REMP 
        // if ((this.state.DateDebut !== null) && (this.state.DateFin !== null) && (this.state.replacedBy.length > 0) && (this.state.numberOfVacationDays > 0) && (this.calculerVacationDaysForSubmit(this.state.numberOfVacationDays,this.state.vacationDaysOfCurrentUser))){
        //   test = false
        // }
        if ((this.state.DateDebut !== null) && (this.state.DateFin !== null) && (this.state.numberOfVacationDays > 0) && (this.calculerVacationDaysForSubmit(this.state.numberOfVacationDays,this.state.vacationDaysOfCurrentUser))){
          test = false
        }
      break; 
    };

    return test
  }

  // calculate the number of vacation days after the soustract 
  public calculerVacationDaysForSubmit = (numberOfVacationRequest, numberOfDaysCurrentUser) => {
    if ((numberOfDaysCurrentUser - numberOfVacationRequest) < 0){
      return false ;
    }
    return true ;
  }


  // Send update of vacation days after send the request
  public updateNumberOfVacationDaysForUser = async (numberOfVacationRequest, numberOfDaysCurrentUser, itemID) => {
    const result = numberOfDaysCurrentUser - numberOfVacationRequest
    const list = sp.web.lists.getByTitle("usersVacationDays");

    const i = await list.items.getById(itemID).update({
      NumberOfDays: result,
    });
  }



  // initialise disable end date for default vacation types
  public ChangeDefaultEndDateDisabled = (beginDate) => {
    if (this.state.disabledDays){
      var defaultNumberOfVacationDay, endDate
      switch (this.state.motifAbsence) {
        case this.state.DemiJournee:
          defaultNumberOfVacationDay = 0.5;
          endDate = beginDate
          break;

        case this.state.Naissance:
          defaultNumberOfVacationDay = 2;
          endDate = addDays(beginDate,defaultNumberOfVacationDay)
        break;

        case this.state.Circonsion:
          defaultNumberOfVacationDay = 1;
          endDate = addDays(beginDate,defaultNumberOfVacationDay)
        break;

        case this.state.Mariage:
          if (this.state.mariageOtionData !== ""){
            defaultNumberOfVacationDay = this.state.numberOfVacationDays;
            endDate = addDays(beginDate,defaultNumberOfVacationDay)
          }
          
        break;

        case this.state.Deces:
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
      if (this.state.motifAbsence === this.state.Congepaye || this.state.motifAbsence === this.state.Maladie) {
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
      const numberOfDays = this.getNumberOfDays(this.state.DateDebut, newDateFormat)

      this.setState({DateFin:newDateFormat, numberOfVacationDays:numberOfDays});
    }
      
  }



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



  // get new user when we change a new user in "Remplac?? Par" 
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




  // Setstate the default of vacation days if the user select "D??c??s"
  public defaultVacationDays = () => {
    var defaultNumberOfVacationDay = 0
    var disabledDays = false
    if (this.state.motifAbsence !== ""){
      switch (this.state.motifAbsence) {
        case this.state.DemiJournee:
          defaultNumberOfVacationDay = 0.5;
          disabledDays = true
          break;
        case this.state.Naissance:
          defaultNumberOfVacationDay = 2;
          disabledDays = true
        break;
        case this.state.Circonsion:
          defaultNumberOfVacationDay = 1;
          disabledDays = true
        break;
      }
    }
    this.setState({DateDebut:null, DateFin:null, disabledDays:disabledDays, numberOfVacationDays:defaultNumberOfVacationDay})
  }





  // Setstate the default of vacation days if the user select "D??c??s"
  public defaultVacationDaysMariage = () => {
    var defaultNumberOfVacationDay = 0
    if (this.state.mariageOtionData !== ""){
      switch (this.state.mariageOtionData) {
        case this.state.Mariage:
          defaultNumberOfVacationDay = 3;
          break;
        case this.state.MariageEnfant:
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
        case this.state.Parents:
          defaultNumberOfVacationDay = 3;
          break;
        case this.state.Conjoint:
          defaultNumberOfVacationDay = 3;
        break;
        case this.state.Enfants:
          defaultNumberOfVacationDay = 3;
        break;

        case this.state.GrandParent:
          defaultNumberOfVacationDay = 2;
        break;
        case this.state.Freres:
          defaultNumberOfVacationDay = 2;
        break;
        case this.state.Soeurs:
          defaultNumberOfVacationDay = 2;
        break;
        case this.state.PetitEnfants:
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

    if (this.state.motifAbsence === this.state.Congepaye || this.state.motifAbsence === this.state.DemiJournee  || this.state.motifAbsence === this.state.Naissance || this.state.motifAbsence === this.state.Circonsion ){
      var replacedByID = 0
      if (this.state.replacedBy.length !== 0) replacedByID = this.state.replacedBy[0].ID
      formData = {
        'Comment': this.state.comment,
        'EndDate': this.state.DateFin,
        'DetailMotifAbsence': "",
        'NrbDays': this.state.numberOfVacationDays,
        'RemainingDays': "0",
        'ReplacedById': replacedByID,
        'ReplacedByStringId': replacedByID.toString(),
        'RequestType': "en cours",
        'ctgVacation': this.state.motifAbsence,
        'dateDeDepart': this.state.DateDebut,
        'vacationType': "",
      };
    }else {
      var replacedByID = 0
      if (this.state.replacedBy.length !== 0) replacedByID = this.state.replacedBy[0].ID
      formData = {
        'Comment': this.state.comment,
        'EndDate': this.state.DateFin,
        'DetailMotifAbsence': this.state.motifAbsence +" "+detailAbsence,
        'NrbDays': this.state.numberOfVacationDays,
        'RemainingDays': "0",
        'ReplacedById': replacedByID,
        'ReplacedByStringId': replacedByID.toString(),
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
    }
    this.updateNumberOfVacationDaysForUser(this.state.numberOfVacationDays,this.state.vacationDaysOfCurrentUser, this.state.itemID)

    this.setState({alertShowed:true})
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

    switch (languageSelectedID) {    
      // Frensh language
      case 1:
        this.setState({
          Congepaye: languages["frenshOptions"]['Cong?? pay??'],
          DemiJournee: languages["frenshOptions"]['Demi journ??e'],
          Maladie: languages["frenshOptions"]['Maladie'],
          Naissance: languages["frenshOptions"]['Naissance'],
          Mariage: languages["frenshOptions"]['Mariage'],
          Deces: languages["frenshOptions"]['D??c??s'],
          Circonsion: languages["frenshOptions"]['Circoncision'],
          Parents: languages["frenshOptions"]['Parents'],
          Conjoint: languages["frenshOptions"]['Conjoint'],
          Enfants: languages["frenshOptions"]['Enfants'],
          GrandParent: languages["frenshOptions"]['Grands-parents'],
          Freres: languages["frenshOptions"]['Fr??res'],
          Soeurs: languages["frenshOptions"]['S??urs'],
          PetitEnfants: languages["frenshOptions"]['Petits-enfants'],
          MariageEnfant: languages["frenshOptions"]['Mariage d???un enfant'],

          TitreDuPage: languages["frenshOptions"]['TitreDuPage'],
          EmployeeName: languages["frenshOptions"]['EmployeeName'],
          EmailOrganisation: languages["frenshOptions"]['EmailOrganisation'],
          EmployeeID: languages["frenshOptions"]['EmployeeID'],
          Email: languages["frenshOptions"]['Email'],
          Champs: languages["frenshOptions"]['Champs'],
          Reason1: languages["frenshOptions"]['Reason1'],
          Reason2: languages["frenshOptions"]['Reason2'],
          StartDate: languages["frenshOptions"]['StartDate'],
          EndDate: languages["frenshOptions"]['EndDate'],
          Attach: languages["frenshOptions"]['Attach'],
          Jours: languages["frenshOptions"]['Jours'],
          RemplacePar: languages["frenshOptions"]['RemplacePar'],
          CommentFile: languages["frenshOptions"]['Comment'],
          OtherDetails: languages["frenshOptions"]['Other details'],
          Solde: languages["frenshOptions"]['Solde'],
          Enregistrer: languages["frenshOptions"]['Enregistrer'],
          DateFile: languages["frenshOptions"]['Date'],
          directionFile: "ltr",
          ChoisirFichier: languages["frenshOptions"]['ChoisirFichier'],
          TitreMessageValidation: languages["frenshOptions"]['TitreMessageValidation'],
          TextMessageValidation: languages["frenshOptions"]['TextMessageValidation'],
        })
      break;

      // Arabic Language
      case 2:
        this.setState({
          Congepaye: languages["arabeOptions"]['Cong?? pay??'],
          DemiJournee: languages["arabeOptions"]['Demi journ??e'],
          Maladie: languages["arabeOptions"]['Maladie'],
          Naissance: languages["arabeOptions"]['Naissance'],
          Mariage: languages["arabeOptions"]['Mariage'],
          Deces: languages["arabeOptions"]['D??c??s'],
          Circonsion: languages["arabeOptions"]['Circoncision'],
          Parents: languages["arabeOptions"]['Parents'],
          Conjoint: languages["arabeOptions"]['Conjoint'],
          Enfants: languages["arabeOptions"]['Enfants'],
          GrandParent: languages["arabeOptions"]['Grands-parents'],
          Freres: languages["arabeOptions"]['Fr??res'],
          Soeurs: languages["arabeOptions"]['S??urs'],
          PetitEnfants: languages["arabeOptions"]['Petits-enfants'],
          MariageEnfant: languages["arabeOptions"]['Mariage d???un enfant'],

          TitreDuPage: languages["arabeOptions"]['TitreDuPage'],
          EmployeeName: languages["arabeOptions"]['EmployeeName'],
          EmailOrganisation: languages["arabeOptions"]['EmailOrganisation'],
          EmployeeID: languages["arabeOptions"]['EmployeeID'],
          Email: languages["arabeOptions"]['Email'],
          Champs: languages["arabeOptions"]['Champs'],
          Reason1: languages["arabeOptions"]['Reason1'],
          Reason2: languages["arabeOptions"]['Reason2'],
          StartDate: languages["arabeOptions"]['StartDate'],
          EndDate: languages["arabeOptions"]['EndDate'],
          Attach: languages["arabeOptions"]['Attach'],
          Jours: languages["arabeOptions"]['Jours'],
          RemplacePar: languages["arabeOptions"]['RemplacePar'],
          CommentFile: languages["arabeOptions"]['Comment'],
          OtherDetails: languages["arabeOptions"]['Other details'],
          Solde: languages["arabeOptions"]['Solde'],
          Enregistrer: languages["arabeOptions"]['Enregistrer'],
          DateFile: languages["arabeOptions"]['Date'],
          directionFile: "rtl",
          ChoisirFichier: languages["arabeOptions"]['ChoisirFichier'],
          TitreMessageValidation: languages["arabeOptions"]['TitreMessageValidation'],
          TextMessageValidation: languages["arabeOptions"]['TextMessageValidation'],

        })
      break;
      
      // English Language
      case 3:
        this.setState({
          Congepaye: languages["englishOptions"]['Cong?? pay??'],
          DemiJournee: languages["englishOptions"]['Demi journ??e'],
          Maladie: languages["englishOptions"]['Maladie'],
          Naissance: languages["englishOptions"]['Naissance'],
          Mariage: languages["englishOptions"]['Mariage'],
          Deces: languages["englishOptions"]['D??c??s'],
          Circonsion: languages["englishOptions"]['Circoncision'],
          Parents: languages["englishOptions"]['Parents'],
          Conjoint: languages["englishOptions"]['Conjoint'],
          Enfants: languages["englishOptions"]['Enfants'],
          GrandParent: languages["englishOptions"]['Grands-parents'],
          Freres: languages["englishOptions"]['Fr??res'],
          Soeurs: languages["englishOptions"]['S??urs'],
          PetitEnfants: languages["englishOptions"]['Petits-enfants'],
          MariageEnfant: languages["englishOptions"]['Mariage d???un enfant'],

          TitreDuPage: languages["englishOptions"]['TitreDuPage'],
          EmployeeName: languages["englishOptions"]['EmployeeName'],
          EmailOrganisation: languages["englishOptions"]['EmailOrganisation'],
          EmployeeID: languages["englishOptions"]['EmployeeID'],
          Email: languages["englishOptions"]['Email'],
          Champs: languages["englishOptions"]['Champs'],
          Reason1: languages["englishOptions"]['Reason1'],
          Reason2: languages["englishOptions"]['Reason2'],
          StartDate: languages["englishOptions"]['StartDate'],
          EndDate: languages["englishOptions"]['EndDate'],
          Attach: languages["englishOptions"]['Attach'],
          Jours: languages["englishOptions"]['Jours'],
          RemplacePar: languages["englishOptions"]['RemplacePar'],
          CommentFile: languages["englishOptions"]['Comment'],
          OtherDetails: languages["englishOptions"]['Other details'],
          Solde: languages["englishOptions"]['Solde'],
          Enregistrer: languages["englishOptions"]['Enregistrer'],
          DateFile: languages["englishOptions"]['Date'],
          directionFile: "ltr",
          ChoisirFichier: languages["englishOptions"]['ChoisirFichier'],
          TitreMessageValidation: languages["englishOptions"]['TitreMessageValidation'],
          TextMessageValidation: languages["englishOptions"]['TextMessageValidation'],

        })

      break;
      default:
        this.setState({
          Congepaye: languages["englishOptions"]['Cong?? pay??'],
          DemiJournee: languages["englishOptions"]['Demi journ??e'],
          Maladie: languages["englishOptions"]['Maladie'],
          Naissance: languages["englishOptions"]['Naissance'],
          Mariage: languages["englishOptions"]['Mariage'],
          Deces: languages["englishOptions"]['D??c??s'],
          Circonsion: languages["englishOptions"]['Circoncision'],
          Parents: languages["englishOptions"]['Parents'],
          Conjoint: languages["englishOptions"]['Conjoint'],
          Enfants: languages["englishOptions"]['Enfants'],
          GrandParent: languages["englishOptions"]['Grands-parents'],
          Freres: languages["englishOptions"]['Fr??res'],
          Soeurs: languages["englishOptions"]['S??urs'],
          PetitEnfants: languages["englishOptions"]['Petits-enfants'],
          MariageEnfant: languages["englishOptions"]['Mariage d???un enfant'],

          TitreDuPage: languages["englishOptions"]['TitreDuPage'],
          EmployeeName: languages["englishOptions"]['EmployeeName'],
          EmailOrganisation: languages["englishOptions"]['EmailOrganisation'],
          EmployeeID: languages["englishOptions"]['EmployeeID'],
          Email: languages["englishOptions"]['Email'],
          Champs: languages["englishOptions"]['Champs'],
          Reason1: languages["englishOptions"]['Reason1'],
          Reason2: languages["englishOptions"]['Reason2'],
          StartDate: languages["englishOptions"]['StartDate'],
          EndDate: languages["englishOptions"]['EndDate'],
          Attach: languages["englishOptions"]['Attach'],
          Jours: languages["englishOptions"]['Jours'],
          RemplacePar: languages["englishOptions"]['RemplacePar'],
          CommentFile: languages["englishOptions"]['Comment'],
          OtherDetails: languages["englishOptions"]['Other details'],
          Solde: languages["englishOptions"]['Solde'],
          Enregistrer: languages["englishOptions"]['Enregistrer'],
          DateFile: languages["englishOptions"]['Date'],
          directionFile: "ltr",
          ChoisirFichier: languages["englishOptions"]['ChoisirFichier'],
          TitreMessageValidation: languages["englishOptions"]['TitreMessageValidation'],
          TextMessageValidation: languages["englishOptions"]['TextMessageValidation'],

        })
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
    
    // options of Absence
    const motifAbsence = [
      {key: this.state.Congepaye,text: this.state.Congepaye},
      {key: this.state.DemiJournee, text: this.state.DemiJournee}, 
      {key: this.state.Maladie, text: this.state.Maladie}, 
      {key: this.state.Naissance, text: this.state.Naissance}, 
      {key: this.state.Mariage, text: this.state.Mariage}, 
      {key: this.state.Deces, text: this.state.Deces}, 
      {key: this.state.Circonsion, text: this.state.Circonsion},
    ];

    // the other options when user choice is "D??c??s"
    const decesOptions = [
      { key: this.state.Parents, text: this.state.Parents, },
      { key: this.state.Conjoint, text: this.state.Conjoint, },
      { key: this.state.Enfants, text: this.state.Enfants, },
      { key: this.state.GrandParent, text: this.state.GrandParent, },
      { key: this.state.Freres, text: this.state.Freres, },
      { key: this.state.Soeurs, text: this.state.Soeurs, },
      { key: this.state.PetitEnfants, text: this.state.PetitEnfants, }
    ];

    // the other options when user choice is "Mariage"
    const mariageOptions =  [
      { key: this.state.Mariage, text: this.state.Mariage, },
      { key: this.state.MariageEnfant, text: this.state.MariageEnfant, }
    ];


    // date picker info
    const DatePickerStrings: IDatePickerStrings = {
      months: ['Janvier', 'F??vrier', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Aout', 'Septembre', 'Octobre', 'Novembre', 'D??cembre'],
      shortMonths: ['Jan', 'Feb', 'Mar', 'Avr', 'Mai', 'Jun', 'Jul', 'Aou', 'Sep', 'Oct', 'Nov', 'Dec'],
      days: ['Diamanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'],
      shortDays: ['DI', 'LU', 'MA', 'ME', 'JE', 'VE', 'SA'],
      goToToday: "Aller ?? aujourd'hui",
      prevMonthAriaLabel: 'Aller au mois pr??c??dent',
      nextMonthAriaLabel: 'Aller au mois prochain',
      prevYearAriaLabel: "Aller ?? l'ann??e pr??c??dente",
      nextYearAriaLabel: "Aller ?? l'ann??e prochaine",
      invalidInputErrorMessage: 'Invalid date format.'
    };

    // get the current date 
    const CurrentDate = new Date().getDate().toString() + "/" + (new Date().getMonth()+1).toString() + "/" + new Date().getFullYear().toString();
    

    const theme = getTheme();       // get the theme of sharepoint
    this.getNumberOfVacationDays()  // Get vacation days number of current user
    // this.getHolidayDays();       // get holidays days of year in current country
    

    return (
      <Fabric
        className="App"
        style={{ background: theme.semanticColors.bodyBackground, color: theme.semanticColors.bodyText }}
      >
        <div className={stylescustom.vacationRequest} dir={this.state.directionFile}>
          <div className={stylescustom.DC}>
            <p className={stylescustom.datenow}>{this.state.DateFile} : <span className="date-time">{CurrentDate}</span></p>
            {/* <div className={stylescustom.titleh1}>Demande de cong?? </div> */}
            <div className={stylescustom.titleh1}>{this.state.TitreDuPage}</div>
            <div className={stylescustom.line}></div>


            


            {/* <p className={stylescustom.indique}>* Indique un champ obligatoire</p> */}
            <div className={stylescustom.row}>

              {/* Select absence Motif */}
              <div className={stylescustom.data}>
                {/* <p className={stylescustom.title}>* Motif d'absence :</p> */}
                <p className={stylescustom.title}>* {this.state.Reason1} :</p>
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



              {/* ********* Show other d??c??s options when user select d??c??s in motif d'absence ********* */}
              {this.state.motifAbsence == this.state.Deces && <div className={stylescustom.data}>
                {/* <p className={stylescustom.title}>* Plus de d??tails</p> */}
                <p className={stylescustom.title}>* {this.state.Reason2}</p>
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
              {this.state.motifAbsence == this.state.Mariage && <div className={stylescustom.data}>
                {/* <p className={stylescustom.title}>* Plus de d??tails</p> */}
                <p className={stylescustom.title}>* {this.state.Reason2}</p>
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
                <p className={stylescustom.title}>* {this.state.StartDate} :</p>
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
                <p className={stylescustom.title}>* {this.state.EndDate}:</p>
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
                  {this.state.motifAbsence === this.state.Maladie && <span>*</span>}{this.state.Attach} :
                </p>
                {/* <label htmlFor="uploadFile" className={stylescustom.btn}>Choisir un ??l??ment</label> */}
                <label htmlFor="uploadFile" className={stylescustom.btn}>{this.state.ChoisirFichier}</label>
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
                <p className={stylescustom.title}>{this.state.Jours} :</p>

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
                {/* <p className={stylescustom.title}>Remplac?? par :</p> */}
                <p className={stylescustom.title}>{this.state.RemplacePar} :</p>
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
                <p className={stylescustom.title}>{this.state.CommentFile} :</p>
                <TextField className={controlClass.TextField} value={this.state.comment} multiline onChange={this.handleChange} />
              </div>
            </div>


            <table className={stylescustom.ad}>
              <thead>
                {/* <th className={stylescustom.title} >Autres d??tails</th> */}
                <th className={stylescustom.title} >{this.state.OtherDetails}</th>
              </thead>
              <tbody className={stylescustom.tbody}>
                <tr>
                  {/* <td className={stylescustom.key}>Solde des cong??s </td> */}
                  <td className={stylescustom.key}>{this.state.Solde} </td>
                  <td className={stylescustom.value}>{this.state.vacationDaysOfCurrentUser}</td>

                </tr>
              </tbody>
            </table>



            <div className={stylescustom.btncont}>
              {/* {this.state.loadingFile ? <Spinner size={SpinnerSize.large} className={stylescustom.spinner} /> : ""} */}
              {/* <button className={stylescustom.btn} onClick={()=>this.collectAllData()} disabled={this.disabledSubmitButton()} >soumettre la demande</button> */}
              <button className={stylescustom.btn} onClick={()=>this.collectAllData()} disabled={this.disabledSubmitButton()} >{this.state.Enregistrer}</button>
            </div>


            {/* <SweetAlert
            show={this.state.alertShowed} title="Demande de cong??" text="Demande envoy??e"
            confirmButtonColor='#7D2935'
            onConfirm={() => window.open(this.props.url + "/SitePages/Vacation-List.aspx", "_self")}
            imageWidth="200"
            imageHeight="200"
            /> */}
            <SweetAlert
            show={this.state.alertShowed} title={this.state.TitreMessageValidation} text={this.state.TextMessageValidation}
            confirmButtonColor='#7D2935'
            onConfirm={() => window.open(this.props.url + "/SitePages/Vacation-List.aspx", "_self")}
            imageWidth="200"
            imageHeight="200"
            />



            {/* <SweetAlert
              show={this.state.alert} title="Demande de cong??" text="Demande envoy??e"
              imageUrl={img}
              confirmButtonColor='#7D2935'
              onConfirm={() => window.open(this.props.webURL + "/SitePages/Tableau-de-bord-utilisateur-des-demandes-de-cong??.aspx", "_self")}
              imageWidth="200"
              imageHeight="200"
            />
            <SweetAlert
              show={this.state.alerteligibility} title="Demande de cong??" text="Votre solde de cong?? est insuffisant"
              imageUrl={img}
              confirmButtonColor='#7D2935'
              onConfirm={() => this.setState({ alerteligibility: false })}
              // onConfirm={() => window.open(this.props.webURL + "/SitePages/Demande-de-cong??.aspx", "_self")}
              imageWidth="200"
              imageHeight="200"
            /> */}
          </div>
        </div>
      </Fabric>
    );
  }
}
