import * as React from 'react';
// import styles from './ModificationConge.module.scss';
import styles from './ModificationConge.module.scss';
import stylescustom from './ModificationCongeCustom.module.scss';
import { IModificationCongeProps } from './IModificationCongeProps';
import { addDays } from 'date-fns';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { escape } from '@microsoft/sp-lodash-subset';
import { Dropdown, IDropdownOption, IDropdownProps, IDropdownStyles } from 'office-ui-fabric-react/lib/Dropdown';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { DatePicker, IDatePickerStrings, IDatePickerStyles, IDatePickerStyleProps } from 'office-ui-fabric-react/lib/DatePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { sp, Web, IWeb, IAttachmentInfo } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItem } from "@pnp/sp/items/types";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import "@pnp/sp/items";
import "@pnp/sp/attachments";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { getTheme } from "@uifabric/styling";
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
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { MSGraphClient, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IUser {
  displayName: string;
  mail:string;
}

export default class ModificationConge extends React.Component<IModificationCongeProps, {}> {
  // state of webPart
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
    replacedByLoginName: "", 
    numberOfVacationDays: 0,  // number of days in vacation
    numberOfVacationDaysForUser: 0,
    ////////////////////////////////////////////////////////////////////////

    ////////// Enabled and disabled buttons for dynamic form ///////////////
    disabledDays: false,   // disable number of days
    disableSubmitButton: true,  // disable submit button 
    alertShowed: false,  // show sucess alert 
    ////////////////////////////////////////////////////////////////////////

    // ID of vaation request
    itemID: 0,

    // data of vacation request
    data : []
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


  // enable and disable submit button
  private disabledSubmitButton = () => {
    var test = true
    switch (this.state.motifAbsence) {
      // 1
      case "Demi journée":
      case "Naissance":
      case "Circoncision":
        // test DA, remplacer par
        if ((this.state.DateDebut !== undefined)&&(this.state.replacedBy.length > 0)){
          test = false
        }
      break;

      // 2
      case "Mariage":
      case "Décès":
        // test Décès options
        // test DA, REMP, Plus details
        if (this.state.motifAbsence === "Décès"){
          if ((this.state.DateDebut !== undefined) && (this.state.replacedBy.length > 0) && (this.state.decesOptionData !== "")){
            test = false
          }
        }
        // test Mariage options
        // test DA, REMP, Plus details
        if (this.state.motifAbsence === "Mariage"){
          if ((this.state.DateDebut !== undefined) && (this.state.replacedBy.length > 0) && (this.state.mariageOtionData !== "")){
            test = false
          }

        }
      break;
      
      // 3
      case "Maladie":
        // test DA, DF, REMP, File
        if ((this.state.DateDebut !== undefined) && (this.state.DateFin !== undefined) && (this.state.replacedBy.length > 0) && (this.state.fileName !== "") && (this.state.numberOfVacationDays > 0)){
          test = false
        }
      break;
      
      // 4
      case "Congé payé":
        // tester DA, DF, REMP 
        if ((this.state.DateDebut !== undefined) && (this.state.DateFin !== undefined) && (this.state.replacedBy.length > 0) && (this.state.numberOfVacationDays > 0)){
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
        case 'Demi journée':
          defaultNumberOfVacationDay = 0.5;
          endDate = beginDate
          break;

        case 'Naissance':
          defaultNumberOfVacationDay = 2;
          endDate = addDays(beginDate,defaultNumberOfVacationDay)
        break;

        case 'Circoncision':
          defaultNumberOfVacationDay = 1;
          endDate = addDays(beginDate,defaultNumberOfVacationDay)
        break;

        case 'Mariage':
          if (this.state.mariageOtionData !== ""){
            defaultNumberOfVacationDay = this.state.numberOfVacationDays;
            endDate = addDays(beginDate,defaultNumberOfVacationDay)
          }
          
        break;

        case 'Décès':
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
      this.setState({DateDebut:newDateFormat, DateFin:null, numberOfVacationDays:0});
      this.ChangeDefaultEndDateDisabled(newDateFormat)

    // if beginDate param equal to false -> this function setstate the endDate
    }else {
      // convert the type of date
      const date = Year.toString() + "-" + Month.toString() + "-" + Day.toString()  + " GMT";
      
      const newDateFormat = new Date(date);
      const diffDays = this.SumVacationDays(this.state.DateDebut, newDateFormat)
      this.setState({DateFin:newDateFormat, numberOfVacationDays:diffDays});
    }
  }


  public SumVacationDays = (dateDebut, DateFin) => {
    var diffDays = DateFin.getTime() - dateDebut.getTime();
    diffDays = diffDays / (1000 * 3600 * 24);
    return diffDays
  }




  // Condition for disable endDate if user select "Demi Journée","Naissance","Mariage","Décès" or "Circoncision"
  private disableEndDate = () => {
    if ((this.state.motifAbsence === "Demi journée") || (this.state.motifAbsence === 'Naissance') || (this.state.motifAbsence === 'Mariage') || ((this.state.motifAbsence === 'Décès')) || ((this.state.motifAbsence === 'Circoncision'))) {
      return true;
    }
    return false;
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



  // get new user when we change a new user in "Remplacé Par" 
  public _getPeoplePickerItems = async (items: any[]) => {
    if (items.length > 0) {
      if (items[0].id && items[0].text && items[0].secondaryText){
        let replacedUserData = {ID:items[0].id,name:items[0].text,email:items[0].secondaryText}
        console.log(replacedUserData)
        this.setState({replacedBy:[replacedUserData]})
      }
    }else {
      console.log('test')
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
        case 'Demi journée':
          defaultNumberOfVacationDay = 0.5;
          disabledDays = true
          break;
        case 'Naissance':
          defaultNumberOfVacationDay = 2;
          disabledDays = true
        break;
        case 'Circoncision':
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
        case 'Mariage':
          defaultNumberOfVacationDay = 3;
          break;
        case 'Mariage d’un enfant':
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
        case 'Conjoint':
          defaultNumberOfVacationDay = 3;
        break;
        case 'Enfants':
          defaultNumberOfVacationDay = 3;
        break;

        case 'Grands-parents':
          defaultNumberOfVacationDay = 2;
        break;
        case 'Frères':
          defaultNumberOfVacationDay = 2;
        break;
        case 'Sœurs':
          defaultNumberOfVacationDay = 2;
        break;
        case 'Petits-enfants':
          defaultNumberOfVacationDay = 2;
        break; 
      }
      this.setState({disabledDays:true, DateDebut:null, DateFin:null})
    }
    this.setState({numberOfVacationDays:defaultNumberOfVacationDay})
  }




  // Get vacation data with current ID
  public GetFormData = async () => {

    // Get ID of vacation list from URL
    var myParm = parseInt(new UrlQueryParameterCollection(window.location.href).getValue("itemId"))
    var data = await Web(this.props.absoluteUrl).lists.getByTitle("vacationRequest").items.getById(myParm)();


    // get all attachments of vacation request selected
    var fileName = ""
    const info: IAttachmentInfo[] = await Web(this.props.absoluteUrl).lists.getByTitle("vacationRequest").items.getById(myParm).attachmentFiles();
    if (info.length !== 0) fileName = info[0].FileName

    const user = await Web(this.props.absoluteUrl).getUserById(data.ReplacedById)();
    switch (data.ctgVacation) {
      case 'Congé payé':
        this.setState({
          comment: data.Comment ,
          DateFin: new Date(data.EndDate) ,
          numberOfVacationDays: data.NrbDays ,
          motifAbsence: data.ctgVacation ,
          DateDebut: new Date(data.dateDeDepart) ,
          fileName: fileName ,
          replacedBy : [{
            text: user.Title ,
            secondaryText: user.UserPrincipalName ,
            loginName: user.LoginName ,
            id: user.Id
          }],
          replacedByLoginName: user.UserPrincipalName
        })
        break;

      case 'Demi journée':
        this.setState({
          comment: data.Comment ,
          DateFin: new Date(data.EndDate) ,
          numberOfVacationDays: data.NrbDays ,
          motifAbsence: data.ctgVacation,
          DateDebut: new Date(data.dateDeDepart),
          fileName: fileName ,
          replacedBy : [{
            text: user.Title ,
            secondaryText: user.UserPrincipalName ,
            loginName: user.LoginName ,
            id: user.Id
          }],
          replacedByLoginName: user.UserPrincipalName
        })
        break;

      case 'Maladie':
        this.setState({
          comment: data.Comment ,
          DateFin: new Date(data.EndDate) ,
          numberOfVacationDays: data.NrbDays ,
          motifAbsence: data.ctgVacation,
          DateDebut: new Date(data.dateDeDepart),
          fileName: fileName ,
          replacedBy : [{
            text: user.Title ,
            secondaryText: user.UserPrincipalName ,
            loginName: user.LoginName ,
            id: user.Id
          }],
          replacedByLoginName: user.UserPrincipalName
        })
        break;

      case 'Naissance':
        this.setState({
          comment: data.Comment ,
          DateFin: new Date(data.EndDate) ,
          numberOfVacationDays: data.NrbDays ,
          motifAbsence: data.ctgVacation,
          DateDebut: new Date(data.dateDeDepart),
          fileName: fileName ,
          replacedBy : [{
            text: user.Title ,
            secondaryText: user.UserPrincipalName ,
            loginName: user.LoginName ,
            id: user.Id
          }],
          replacedByLoginName: user.UserPrincipalName
        })
        break;

      case 'Mariage':
        this.setState({
          comment: data.Comment ,
          DateFin: new Date(data.EndDate) ,
          numberOfVacationDays: data.NrbDays ,
          motifAbsence: data.ctgVacation,
          DateDebut: new Date(data.dateDeDepart),
          fileName: fileName ,
          mariageOtionData : data.vacationType,
          replacedBy : [{
            text: user.Title ,
            secondaryText: user.UserPrincipalName ,
            loginName: user.LoginName ,
            id: user.Id
          }],
          replacedByLoginName: user.UserPrincipalName
        })
        break;

      case 'Décès':
        this.setState({
          comment: data.Comment ,
          DateFin: new Date(data.EndDate) ,
          numberOfVacationDays: data.NrbDays ,
          motifAbsence: data.ctgVacation,
          DateDebut: new Date(data.dateDeDepart),
          fileName: fileName ,
          decesOptionData : data.vacationType,
          replacedBy : [{
            text: user.Title ,
            secondaryText: user.UserPrincipalName ,
            loginName: user.LoginName ,
            id: user.Id
          }],
          replacedByLoginName: user.UserPrincipalName
        })
        break;

      case 'Circoncision':
        this.setState({
          comment: data.Comment ,
          DateFin: new Date(data.EndDate) ,
          numberOfVacationDays: data.NrbDays ,
          motifAbsence: data.ctgVacation,
          DateDebut: new Date(data.dateDeDepart),
          fileName: fileName ,
          replacedBy : [{
            text: user.Title ,
            secondaryText: user.UserPrincipalName ,
            loginName: user.LoginName ,
            id: user.Id
          }],
          replacedByLoginName: user.UserPrincipalName
        })
        break;
    }
  }



  public sendFormData = async () => {
    if (this.state.disableSubmitButton !== false){
      
      // initialise data
      var detailAbsence = ""
      var formData
      if(this.state.decesOptionData !== "") detailAbsence = this.state.decesOptionData
      if(this.state.mariageOtionData !== "") detailAbsence = this.state.mariageOtionData

      // if vacation request with motifAbsence like ("Congé payé", "Demi journée", "Naissance", "Circoncision")
      if (this.state.motifAbsence === "Congé payé" || this.state.motifAbsence === "Demi journée"  || this.state.motifAbsence === "Naissance" || this.state.motifAbsence === "Circoncision" ){
        console.log(this.state.replacedBy[0].ID)
        formData = {
          'Comment': this.state.comment,
          'EndDate': this.state.DateFin,
          'DetailMotifAbsence': "",
          'NrbDays': this.state.numberOfVacationDays,
          'RemainingDays': "0",
          'ReplacedById': this.state.replacedBy[0].ID,
          'ReplacedByStringId': this.state.replacedBy[0].ID.toString() ,
          'RequestType': "en cours",
          'ctgVacation': this.state.motifAbsence,
          'dateDeDepart': this.state.DateDebut,
          'vacationType': "",
        }
        // if vacation request with motifAbsence like Mariage or Décès
      }else {
        formData = {
          'Comment': this.state.comment,
          'EndDate': this.state.DateFin,
          'DetailMotifAbsence': this.state.motifAbsence +" "+detailAbsence,
          'NrbDays': this.state.numberOfVacationDays,
          'RemainingDays': "0",
          'ReplacedById': this.state.replacedBy[0].id,
          'ReplacedByStringId': this.state.replacedBy[0].id.toString() ,
          'RequestType': "en cours",
          'ctgVacation': this.state.motifAbsence,
          'dateDeDepart': this.state.DateDebut,
          'vacationType': detailAbsence,
        }
      }

      // update selected vacation request 

      var myParm = parseInt(new UrlQueryParameterCollection(window.location.href).getValue("itemId"))
      console.log(myParm)
      const sendData = await Web(this.props.absoluteUrl).lists.getByTitle("vacationRequest").items.getById(myParm).update(formData)
      this.setState({alertShowed:true})


    }
  }



  // Get numbers of vacation days for current User
  public numberOfVacationDays = async() => {
    if (this.state.currentUserID !== ""){
      const items = await Web(this.props.absoluteUrl).lists.getByTitle('usersVacationDays').items();
      var currentUserVacationDays = items.filter(item => this.state.currentUserID === item.ID_user);
      this.setState({numberOfVacationDaysForUser:currentUserVacationDays[0].number_of_vacation})
    }

  }

  componentDidMount(): void {
    this.GetFormData();
    this.getUsers();
  }


  public render(): React.ReactElement<IModificationCongeProps> {

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
      {key: "Congé payé",text: "Congé payé"},
      {key: "Demi journée", text: "Demi journée"}, 
      {key: "Maladie", text: "Maladie"}, 
      {key: "Naissance", text: "Naissance"}, 
      {key: "Mariage", text: "Mariage"}, 
      {key: "Décès", text: "Décès"}, 
      {key: "Circoncision", text: "Circoncision"},
    ];


    // the other options when user choice is "Décès"
    const decesOptions = [
      { key: "Parents", text: "Parents", },
      { key: "Conjoint", text: "Conjoint", },
      { key: "Enfants", text: "Enfants", },
      { key: "Grands-parents", text: "Grands-parents", },
      { key: "Frères", text: "Frères", },
      { key: "Sœurs", text: "Sœurs", },
      { key: "Petits-enfants", text: "Petits-enfants", }
    ];


    // the other options when user choice is "Mariage"
    const mariageOptions =  [
      { key: "Mariage", text: "Mariage", },
      { key: "Mariage d’un enfant", text: "Mariage d’un enfant", }
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

    this.numberOfVacationDays();


    return (
      <Fabric
        className="App"
        style={{ background: theme.semanticColors.bodyBackground, color: theme.semanticColors.bodyText }}
      >
        <div className={stylescustom.vacationRequest}>
          <div className={stylescustom.DC}>
            <p className={stylescustom.datenow}>Date : <span className="date-time">{CurrentDate}</span></p>
            <div className={stylescustom.titleh1}>Modification demande de congé </div>
            <div className={stylescustom.line}></div>


            <div className={stylescustom.row}>
              <div className={stylescustom.col}>
                <table className={stylescustom.table}>
                  <tbody>
                    <tr>
                      <td className={stylescustom.key}>Nom de l'employé</td>
                      <td className={stylescustom.value}>{this.state.currentUserDisplayName} </td>
                    </tr>
                    <tr>
                      <td className={stylescustom.key}>Adresse email de l'organisation</td>
                      <td className={stylescustom.value}>{this.state.currentUserPrincipalName}</td>
                    </tr>
                    <tr>
                      <td className={stylescustom.key}>ID employé</td>
                      <td className={stylescustom.value}>{this.state.currentUserID}</td>
                    </tr>
                    <tr>
                      <td className={stylescustom.key}>Adresse email</td>
                      <td className={stylescustom.value}>{this.state.currentUserMail}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>


            <p className={stylescustom.indique}>* Indique un champ obligatoire</p>
            <div className={stylescustom.row}>

              {/* Select absence Motif */}
              <div className={stylescustom.data}>
                <p className={stylescustom.title}>* Motif d'absence :</p>
                <Dropdown
                  styles={dropdownStyles}
                  options={motifAbsence}
                  onChanged={(value) => this.setState({ motifAbsence:value['key'], disabledDays:false},this.defaultVacationDays)}
                  defaultSelectedKey={this.state.motifAbsence}
                />
              </div>



              {/* ********* Show other décès options when user select décès in motif d'absence ********* */}
              {this.state.motifAbsence == 'Décès' && <div className={stylescustom.data}>
                <p className={stylescustom.title}>* Plus de détails</p>
                <Dropdown
                  styles={dropdownStyles}
                  options={decesOptions}
                  onChanged={(value) => this.setState({decesOptionData:value['key']}, this.defaultVacationDaysDeces)}
                  defaultSelectedKey={this.state.decesOptionData}
                  // errorMessage={this.state.errors.VacationType !== "" ? this.state.errors.VacationType : ""}
                />
              </div>}
              {/* ******************************************************************************************* */}




              {/* ********* Show other mariage options when user select mariage in motif d'absence ********* */}
              {this.state.motifAbsence == 'Mariage' && <div className={stylescustom.data}>
                <p className={stylescustom.title}>* Plus de détails</p>
                <Dropdown
                  styles={dropdownStyles}
                  options={mariageOptions}
                  onChanged={(value) => this.setState({mariageOtionData:value['key']},this.defaultVacationDaysMariage)}
                  defaultSelectedKey={this.state.mariageOtionData}
                  // errorMessage={this.state.errors.VacationType !== "" ? this.state.errors.VacationType : ""}
                />
              </div>}
              {/* ******************************************************************************************* */}


              <div className={stylescustom.data}>
                <p className={stylescustom.title}>* Date debut :</p>
                <DatePicker
                  className={controlClass.TextField}
                  allowTextInput={false}
                  strings={DatePickerStrings}
                  value={this.state.DateDebut}
                  onSelectDate={(e) => { this.SelectDate(e.getDate(), e.getMonth()+1 ,e.getFullYear(), true) }}
                  ariaLabel="Select a date"
                />
              </div>



              <div className={stylescustom.data}>
                <p className={stylescustom.title}>* Date Fin :</p>
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
                  {this.state.motifAbsence === "Maladie" && <span>*</span>}Attacher un élément justificatif :
                </p>
                <label htmlFor="uploadFile" className={stylescustom.btn}>Choisir un élément</label>
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
                <p className={stylescustom.title}>Jours :</p>

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
                <p className={stylescustom.title}>Remplacé par :</p>
                <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={1}
                  required={false}
                  onChange={this._getPeoplePickerItems}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  ensureUser={true}
                  defaultSelectedUsers={[this.state.replacedByLoginName]}
                />
                <svg className={stylescustom.iconsearch} xmlns="http://www.w3.org/2000/svg" width="20" height="20.003" viewBox="0 0 20 20.003">
                  <path data-name="Icon awesome-search" d="M19.728,17.294,15.833,13.4a.937.937,0,0,0-.664-.273h-.637a8.122,8.122,0,1,0-1.406,1.406v.637a.937.937,0,0,0,.273.664l3.895,3.895a.934.934,0,0,0,1.324,0l1.106-1.106a.942.942,0,0,0,0-1.328Zm-11.6-4.168a5,5,0,1,1,5-5A5,5,0,0,1,8.126,13.126Z" />
                </svg>
              </div>
            </div>



            <div className={stylescustom.row}>
              <div className={stylescustom.comment}>
                <p className={stylescustom.title}>Commentaire :</p>
                <TextField className={controlClass.TextField} value={this.state.comment} multiline onChange={this.handleChange} />
              </div>
            </div>


            <table className={stylescustom.ad}>
              <thead>
                <th className={stylescustom.title} >Autres détails</th>
              </thead>
              <tbody className={stylescustom.tbody}>
                <tr>
                  <td className={stylescustom.key}>Solde des congés </td>
                  <td className={stylescustom.value}>{this.state.numberOfVacationDaysForUser}</td>
                </tr>
              </tbody>
            </table>



            <div className={stylescustom.btncont}>
              {/* {this.state.loadingFile ? <Spinner size={SpinnerSize.large} className={stylescustom.spinner} /> : ""} */}
              <button className={stylescustom.btn} onClick={()=>this.sendFormData()} disabled={this.disabledSubmitButton()} >soumettre la demande</button>
            </div>



            <SweetAlert
            show={this.state.alertShowed} title="Modification demande de congé" text="Modification envoyée"
            confirmButtonColor='#7D2935'
            onConfirm={() => window.open(this.props.absoluteUrl + "/SitePages/Vacation-List.aspx", "_self")}
            imageWidth="200"
            imageHeight="200"
            />
            {/*
            window.open(this.props.webURL + "/SitePages/Tableau-de-bord-utilisateur-des-demandes-de-congé.aspx", "_self")
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
