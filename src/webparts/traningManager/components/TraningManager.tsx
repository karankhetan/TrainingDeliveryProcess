import * as React from 'react';
import styles from './TraningManager.module.scss';
import { ITraningManagerProps } from './ITraningManagerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { List, ActionButton, Label, FocusZone, FocusZoneDirection, DefaultButton } from 'office-ui-fabric-react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import DatePicker from "react-datepicker";
import 'react-dates/initialize';
import * as moment from 'moment'
import { SingleDatePicker } from 'react-dates';
import 'react-dates/lib/css/_datepicker.css';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import "react-datepicker/dist/react-datepicker.css";
import { Pivot, PivotItem, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import Modal from 'office-ui-fabric-react/lib-es2015/Modal';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
export interface IListItem {
  DateOfTraining: string;
  Id: number;
  OData__ModerationStatus: number;
  Title: string;
}

export interface IstateItems {
  stateItems: IListItem[];
  DisplayItems: IListItem[];
  filtertext: string;
  showModal: boolean;
  startDate: Date;
  Message: string;
  IsMessageDisplay: boolean;
  Approver: string;
  focus:boolean;
  ApprovedDates:any[];
  TextFieldValue:string;
}

export default class TraningManager extends React.Component<ITraningManagerProps, IstateItems> {

  constructor(props: ITraningManagerProps, state: IstateItems) {
    super(props, state);
    this._onFilterChanged = this._onFilterChanged.bind(this);
    this.changeListItems = this.changeListItems.bind(this);
    this.handleChange = this.handleChange.bind(this);

    this.state = {
      stateItems: [],
      DisplayItems: [],
      filtertext: "All Tranings",
      showModal: false,
      startDate: null,//new Date(),
      Message: "",
      IsMessageDisplay: false,
      Approver: "",
      focus:false,
      ApprovedDates:[],
      TextFieldValue:""
    };
  }

  private InitializeListItem(){
    let TempArray: IListItem[] = [];
     this.props.getlistItem().then((res) => {
      res.map((item) => {
        let tempobj: IListItem = { DateOfTraining: item.DateOfTraining, Id: item.Id, OData__ModerationStatus: item.OData__ModerationStatus, Title: item.Title };
        TempArray.push(tempobj);
      });
      this.setState({ stateItems: TempArray });
      
    });
  }

   componentDidMount() {
    let TempArray: IListItem[] = [];
    this.props.getlistItem().then((res) => {
      res.map((item) => {
        let tempobj: IListItem = { DateOfTraining: item.DateOfTraining, Id: item.Id, OData__ModerationStatus: item.OData__ModerationStatus, Title: item.Title };
        TempArray.push(tempobj);
      });
      this.setState({ stateItems: TempArray, DisplayItems: TempArray });
    });
  }
  private _onRenderCell(item: IListItem, index: number | undefined): JSX.Element {
    let currentDate = new Date(item.DateOfTraining);
    return (
      <div data-is-focusable={true} style={{ borderBottom: "1px solid #dadada", position: "relative", minHeight: "30px" }}>
        <span className={styles.Title}>{item.Title}</span>
        <span className={styles.Date}>{currentDate.toDateString()}</span>
        <span className={styles.Status}>{item.OData__ModerationStatus == 0 ? "Approved" : item.OData__ModerationStatus == 1 ? "Rejected" : "Pending"}</span>
        <div className={styles.Revoke}>
          <ActionButton
            style={{ font: "14px", height: "29px" }}
            iconProps={{ iconName: 'Delete' }}
            onClick={this.Delete.bind(this, item)}
          >
          </ActionButton>
        </div>
      </div>
    );
  }

  private _Changed(text: string): void {
    this.setState({TextFieldValue:text});
  }

  public Delete(item: IListItem) {
    const items = this.state.stateItems;
    this.props.DeleteListItem(item.Id).then((res) => {
      this.setState({
        stateItems: item.Id ? items.filter(item1 => item1.Id != item.Id) : items
      }, () => { this._onFilterChanged("") });
    });
  }

  private _onFilterChanged(text: string): void {
    let items = this.state.stateItems;

    if (this.state.filtertext == "All Trainings") { }
    else if (this.state.filtertext == "Approved Trainings") {
      items = items.filter(item => item.OData__ModerationStatus == 0)
    }
    else if (this.state.filtertext == "Rejected Trainings") {
      items = items.filter(item => item.OData__ModerationStatus == 1)
    }
    else {
      items = items.filter(item => item.OData__ModerationStatus == 2)
    }

    this.setState({
      DisplayItems: text ? items.filter(item => item.Title.toLowerCase().indexOf(text.toLowerCase()) >= 0) : items
    });
  }

  public changeListItems(item: PivotItem): void {
    this.InitializeListItem();
    this.setState({ filtertext: item.props.headerText }, () => { this._onFilterChanged("") });
  }
  private _showModal = (): void => {
    let temparray=[];
    this.props.context.spHttpClient.get(`${this.props.context.pageContext.site.absoluteUrl}/_api/web/lists/GetByTitle('TraningEventList')/items?$select=OData__ModerationStatus,DateOfTraining&$filter=OData__ModerationStatus eq 0`, SPHttpClient.configurations.v1).then(
      (spHttpClientResponse: SPHttpClientResponse) => {
        spHttpClientResponse.json().then(
          (jsonresponse: any) => {
            jsonresponse.value.map((mapitem)=>{
              temparray.push(moment(new Date(mapitem.DateOfTraining)));
            });
          }
        );
      }
    );


    this.setState({ showModal: true,ApprovedDates:temparray });
  };

  private _closeModal = (): void => {
    this.setState({ showModal: false });
  };
  handleChange(date) {
    this.setState({
      startDate: date
    });
  }
  private _getPeoplePickerItems(items: any[]) {
    if (items.length > 0) {
      let obj = items[0].secondaryText;
      this.setState({ Approver: obj })
    }
  }

  private _addNewListItem() {
    if(this.state.TextFieldValue=="" || this.state.startDate==null || this.state.Approver.length <=0){
      this.setState({ IsMessageDisplay: true, Message: "Please Enter all the details..." });
    }
    else{
    this.setState({ IsMessageDisplay: true, Message: "Please Wait..." });
    let date= this.state.startDate.toJSON().substring(0,this.state.startDate.toJSON().indexOf("T"));
    const body: string = JSON.stringify({
      'Title': this.state.TextFieldValue,
      'DateOfTraining': date,
      'Approver': this.state.Approver
    });
    console.log(body)
    this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TraningEventList')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then((item: IListItem): void => {
        console.log(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
        let newobj:IListItem={Id:item.Id,DateOfTraining:item.DateOfTraining,OData__ModerationStatus:item.OData__ModerationStatus,Title:item.Title};
        this.setState({stateItems:[...this.state.stateItems,newobj],IsMessageDisplay:false},()=>{this._onFilterChanged("");this._closeModal();});

      }, (error: any): void => {
        console.log('Error while creating the item: ' + error);
      });
    }
  }




  public render(): React.ReactElement<ITraningManagerProps> {
    const modalelement = (<div>
      <DefaultButton onClick={this._showModal} className={styles.button} text="Add New Training Event" />
      <Modal
        isOpen={this.state.showModal}
        onDismiss={this._closeModal}
        isBlocking={true}
        containerClassName={styles.container}
      >
        <div className={styles.header}>
          <span>Add New Training Event</span>
        </div>
        <div className={styles.body} >
          <div>
            <Label>Title :</Label>
            <TextField placeholder="Enter the Title" id="TitleText" onBeforeChange={this._Changed.bind(this)}></TextField>
            <Label>Traning Date (Note: The dates which are blocked are having training already approved):</Label>
            <SingleDatePicker
              date={this.state.startDate} // momentPropTypes.momentObj or null
              onDateChange={date => this.setState({ startDate:date })} // PropTypes.func.isRequired
              focused={this.state.focus} // PropTypes.bool
              onFocusChange={({ focused }) => this.setState({ focus:focused })} // PropTypes.func.isRequired
              id="your_unique_id" // PropTypes.string.isRequired,
              numberOfMonths={2}
              //isDayHighlighted={day1 => this.returnDates().some(day2 => isSameDay(day1, day2))}
              isDayBlocked={day1 =>  this.state.ApprovedDates.some(day2=> day1.isSame(day2,'d'))}
            />
            <Label>Approver :</Label>
            <PeoplePicker
              context={this.props.context}
              //defaultSelectedUsers={[this.props.context.pageContext.user.email]}
              personSelectionLimit={1}
              disabled={false}
              selectedItems={this._getPeoplePickerItems.bind(this)}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              groupName={"approvers"}
              resolveDelay={200}
            />
          </div>
          {this.state.IsMessageDisplay ? <div><Label>{this.state.Message}</Label></div> : ""}
          <div style={{ paddingTop: "20px" }}>
            <DefaultButton className={styles.button} onClick={this._closeModal} text="Cancel" />&nbsp;&nbsp;&nbsp;
            <DefaultButton className={styles.button} onClick={this._addNewListItem.bind(this)} text="Add" />
          </div>
        </div>
      </Modal>
    </div>);


    return (
      <div>
        {modalelement}
        <div>
          <Pivot linkSize={PivotLinkSize.large} onLinkClick={this.changeListItems}>
            <PivotItem headerText="All Trainings"            >
            </PivotItem>
            <PivotItem headerText="Approved Trainings">
            </PivotItem>
            <PivotItem headerText="Rejected Trainings">
            </PivotItem>
            <PivotItem headerText="Pending for Approval">
            </PivotItem>
          </Pivot>
        </div>
        <br></br>
        <FocusZone direction={FocusZoneDirection.vertical}>
          <TextField label={'Filter by Title :'} onBeforeChange={this._onFilterChanged} />
          {this.state.DisplayItems.length > 0 ?
            <div>
              <div data-is-focusable={true} style={{ paddingTop: "10px", borderBottom: "1px solid #dadada", position: "relative", minHeight: "30px" }}>
                <span className={styles.Title}>Title</span>
                <span className={styles.Date}>Date</span>
                <span className={styles.Status}>Status</span>
                <div className={styles.Revoke}>
                  Delete
            </div>
              </div>
              <br></br>
              <List
                items={this.state.DisplayItems}
                onRenderCell={this._onRenderCell.bind(this)}
              >

              </List>
            </div>
            :
            <div style={{ paddingTop: "10px" }} >There are no Training events....</div>
          }
        </FocusZone>

      </div>

    );
  }
}
