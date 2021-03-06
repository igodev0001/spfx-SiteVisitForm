import * as React from 'react';
import styles from './SiteVisitForm.module.scss';
import { ISiteVisitFormProps } from './ISiteVisitFormProps';

import {
  PanelType, TextField, PrimaryButton,
  Toggle, Dropdown, IDropdownOption
} from '@fluentui/react/lib/';
import AddVisitorPanel from './AddVisitorPanel';
import AddFlightPanel from './AddFlightPanel';
import EmployeeCombox from "./EmployeeCombox";
import { v4 as uuid } from 'uuid';
import CountryCode from "./CountryCode";

import { TeachingBubble } from '@fluentui/react/lib/TeachingBubble';
import { DirectionalHint } from '@fluentui/react/lib/Callout';
import { useBoolean } from '@uifabric/react-hooks';
import * as Common from './Common';
import { IVisitorForm, IListItem, IVisitor, IFlight, IEmailObject } from './Common';

import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
  HttpClient,
  HttpClientResponse,
  IHttpClientOptions
} from '@microsoft/sp-http';

const tempid: string = uuid();

export default function SiteVisitForm(props: ISiteVisitFormProps) {

  const childRef = React.useRef(null);
  const VisitorRef = React.useRef(null);
  const FlightRef = React.useRef(null);
  const countryCodeRef = React.useRef(null);
  const CreatedBy = props.context.pageContext.user.email;
  const [isDirtyData, setDirty] = React.useState(false);

  // let qListId = Common.func_GetUrlParameter('ListId');
  // let qFormID = Common.func_GetUrlParameter('FormID');
  let qFunc = Common.func_GetUrlParameter('Func');

  if (qFunc === "" || qFunc === undefined) {
    qFunc = "A";
  }

  const currentweburl = props.context.pageContext.web.absoluteUrl;

  const passToSubComponent: ISiteVisitFormProps = {
    description: "",
    context: props.context,
    __content: ":"
  };

  sessionStorage.setItem(`qFunc`, qFunc);

  const [teachingBubbleVisibleForm, { toggle: toggleTeachingBubbleVisibleForm }] = useBoolean(false);
  const [errorMessage, setErrorMessage] = React.useState('');
  const [webControlFocus, setFocus] = React.useState('#btnSubmit');
  const [EmployeeComboxValue, setEmployeeSelectedValue] = React.useState('');
  const [ContryCodeValue, setCountryCodeOption] = React.useState('');

  const [valueofArea, setAreaValue] = React.useState(Common.__defaultitem);
  const [valueofLocation, setLocation] = React.useState(Common.__defaultitem);
  const [valueofPurpose, setPurpose] = React.useState(Common.__defaultitem);
  const [boolAirportPickup, setAirportPickup] = React.useState(false);
  const [boolCarRental, setCarRental] = React.useState(false);
  const [valueofMobile, setvalueofMobile] = React.useState('');
  const [valueofRoomNeeds, setvalueofRoomNeeds] = React.useState('');
  const [valueofComment, setvalueofComment] = React.useState('');
  const [valueofOtherPurpose, setvalueofOtherPurpose] = React.useState('');
  const [isOthersPurposeShow, setOthersTypeShow] = React.useState(false);
  const [isOfficeLocationShow, setOfficeLocationShow] = React.useState(false);
  const [isCarRentalShow, setCarRentalShow] = React.useState(false);

  const onValueUpdate = (empValue) => {
    setEmployeeSelectedValue(empValue);
  };

  const onCountryCodeValueUpdate = (ccValue) => {
    setCountryCodeOption(ccValue);
  };

  const func_exceptionHandle = (errMessage, errObject) => {
    const err = `Whoops, something went wrong...Unexpected Error.Error code:${errMessage}`;
    setErrorMessage(err);
    toggleTeachingBubbleVisibleForm();

    Common.func_SendExceptionEmail(err + ",errorOject=" + JSON.stringify(errObject), props);
  };

  switch (qFunc) {
    case "A":
      sessionStorage.setItem(`formID`, tempid);
      sessionStorage.setItem(`ListId`, "0");
      break;
    case "E":
    case "V":
      sessionStorage.removeItem(`formID`);
      sessionStorage.removeItem(`ListId`);
      alert(`Sorry, we can't find this request. Please contact system admin.`);
      break;
  }

  const _submit = async (e: React.FormEvent) => {
    e.preventDefault();

    let _visitorForm: IVisitorForm = {
      Title: "",
      Office_x0020_Location: "",
      Trip_x0020_Purpose: "",
      Airport_x0020_Pickup: "",
      Car_x0020_Rental: "",
      Point_x0020_of_x0020_Contact: "",
      Mobile_x0020_Number: "",
      Meeting_x0020_Room_x0020_Needs: "",
      Comment: "",
      Purpose_x0020_Other: "",
      FormID: "",
      Country_x0020_Code: "",
      Id: 0
    };

    if (sessionStorage.getItem(`formID`) === null || sessionStorage.getItem(`formID`) === '') {
      setErrorMessage(`Your session was lost! Please login again!`);
      toggleTeachingBubbleVisibleForm();
      return false;
    }


    let tempkey = sessionStorage.getItem(`formID`) + `_visitor`;

    //console.log(`func_addVisitorDetail.storagevisitorkey=` + storagevisitorkey);
    let _VisitorList: IVisitor[] = [];

    _VisitorList = JSON.parse(sessionStorage.getItem(tempkey));

    if (_VisitorList === null) {
      setErrorMessage(`Please add one visitor at least before your submit the form!`);
      toggleTeachingBubbleVisibleForm();
      return false;
    }

    if (valueofArea.key === "please select") {
      setErrorMessage(`Please select the site you are going to visit!`);
      toggleTeachingBubbleVisibleForm();
      return false;
    }

    if (valueofArea.key === "TW") {
      if (valueofLocation.key.toString() === "please select") {
        setErrorMessage(`Please select office locaion!`);
        toggleTeachingBubbleVisibleForm();
        return false;
      }
    }

    if (valueofPurpose.key === "please select") {
      setErrorMessage(`Please select the purpose!`);
      toggleTeachingBubbleVisibleForm();
      return false;
    }

    if (valueofPurpose.key === "Others") {
      if (valueofOtherPurpose === "" || valueofOtherPurpose === undefined) {
        setErrorMessage(`Please fill in other purpose!`);
        toggleTeachingBubbleVisibleForm();
        return false;
      }
    }

    if (EmployeeComboxValue === "" || EmployeeComboxValue === undefined) {
      setErrorMessage(`Please select visitor point of contact!`);
      toggleTeachingBubbleVisibleForm();
      return false;
    }

    if (ContryCodeValue === "please select") {
      setErrorMessage(`Please select country code!`);
      toggleTeachingBubbleVisibleForm();
      return false;
    }

    if (valueofMobile === "" || valueofMobile === undefined) {
      setErrorMessage(`Please fill in mobile number!`);
      toggleTeachingBubbleVisibleForm();
      return false;
    }

    _visitorForm.Title = valueofArea.key.toString();
    _visitorForm.Trip_x0020_Purpose = valueofPurpose.key.toString();
    _visitorForm.Airport_x0020_Pickup = boolAirportPickup ? "Yes" : "No";
    _visitorForm.Car_x0020_Rental = boolCarRental ? "Yes" : "No";
    _visitorForm.Point_x0020_of_x0020_Contact = EmployeeComboxValue;
    _visitorForm.Mobile_x0020_Number = valueofMobile;
    _visitorForm.Meeting_x0020_Room_x0020_Needs = valueofRoomNeeds;
    _visitorForm.Comment = valueofComment;
    _visitorForm.FormID = tempid;
    _visitorForm.Country_x0020_Code = ContryCodeValue;

    if (valueofArea.key === "TW") {
      _visitorForm.Office_x0020_Location = valueofLocation.key.toString();
    }

    if (valueofPurpose.key === "Others") {
      _visitorForm.Purpose_x0020_Other = valueofOtherPurpose;
    }

    let posturl = `${currentweburl}/_api/web/lists/getByTitle('${Common.ListCollection.OnlineVisitForm}')/items`;

    const options: ISPHttpClientOptions = {
      headers: {},
      body: JSON.stringify(_visitorForm)
    };

    await props.context.spHttpClient.post(posturl, SPHttpClient.configurations.v1, options)
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then((item: IListItem): void => {
        if (item.Id === undefined) {
          throw new Error('E000001-OnlineVisitForm' + JSON.stringify(_visitorForm));
        }
        else {

          _visitorForm.Id = item.Id;
          const promiseP = Common.func_ChangeItemPermission(item.Id, _visitorForm.Title, currentweburl, props);

          Promise.all([promiseP]).then(() => {
            const promiseA = Common.func_AddVisitorDetail(item.Id.toString(), currentweburl, props);
            const promiseB = Common.func_AddFlightDetail(item.Id.toString(), currentweburl, props);
            const promiseC = Common.func_SendEmail("No", Common.ParameterCollection.NewRequestEmailSubject, _visitorForm, CreatedBy, currentweburl, props);

            Promise.all([promiseA, promiseB, promiseC]).then(() => {
              window.location.replace(`${currentweburl}/SitePages/Your-request-was-submitted-successfully.aspx`);
            });
          });
        }
      })
      .catch(x => {
        func_exceptionHandle(x.message, _visitorForm);
      });
  };

  const onAreaChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    if (item.key === "TW") {
      setOfficeLocationShow(true);
    }
    else {
      setOfficeLocationShow(false);
    }

    if (item.key === "US") {
      setCarRentalShow(true);
    }
    else {
      setCarRentalShow(false);
    }

    setAreaValue(item);
  };

  const onOfficeLocationChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    setLocation(item);
  };

  const onPurposeChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {

    if (item.key === "Others") {
      setOthersTypeShow(true);
    }
    else {
      setOthersTypeShow(false);
    }
    setPurpose(item);

  };

  function _onAirportPickupChange(ev: React.MouseEvent<HTMLElement>, checked: boolean) {
    setAirportPickup(checked);
  }

  function _onCarRentalChange(ev: React.MouseEvent<HTMLElement>, checked: boolean) {
    setCarRental(checked);
  }

  const func_setvalueofMobile = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    setvalueofMobile(newText);
  };

  const func_setvalueofRoomNeeds = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    setvalueofRoomNeeds(newText);
  };

  const func_setvalueofComment = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    setvalueofComment(newText);
  };

  const func_setvalueofOtherPurpose = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    setvalueofOtherPurpose(newText);
  };

  return (

    <div className={styles.container}>
      <form onSubmit={_submit}>
        <div className={styles.row} >
          <div className={styles.column} >
            [<span className={styles.required}>*</span>] Mandatory.
          </div>
          <div className={styles.column80p} >

          </div>
        </div>
        <div className={styles.row} >
          <div className={styles.column20p} >
            Which site are you visiting<span className={styles.required}>*</span>
          </div>
          <div className={styles.column80p} >
            <Dropdown
              placeholder="Please select"
              options={Common.optionsofArea}
              styles={Common.dropdownStyles}
              onChange={onAreaChange}
              id="ddlArea"
              selectedKey={valueofArea.key}
            />
          </div>
        </div>
        <div className={isOfficeLocationShow ? styles.row : styles.rownoshow} >
          <div className={styles.column20p} >
            Office Location
        </div>
          <div className={styles.column80p} >
            <Dropdown
              placeholder="Please select"
              options={Common.optionsofOfficeLocation}
              styles={Common.dropdownStyles}
              onChange={onOfficeLocationChange}
              id="ddlOfficeLocation"
              selectedKey={valueofLocation.key}
            />
          </div>
        </div>
        <div className={styles.row} >
          <div className={styles.column20p}>
            Visitor List<span className={styles.required}>*</span>
          </div>
          <div className={styles.column80p}>
            <AddVisitorPanel panelType={PanelType.medium} tempid={tempid} subprops={passToSubComponent} ref={VisitorRef} setDirty={setDirty} />
          </div>
        </div>
        <div className={styles.row} >
          <div className={styles.column20p} >
            Trip Purpose<span className={styles.required}>*</span>
          </div>
          <div className={styles.column80p} >
            <Dropdown
              placeholder="Please select"
              options={Common.optionsofPurpose}
              styles={Common.dropdownStyles}
              onChange={onPurposeChange}
              id="ddlPurpose"
              selectedKey={valueofPurpose.key}
            />
            <TextField
              styles={isOthersPurposeShow ? Common.textStylesofNameBlock : Common.textStylesofName}
              value={valueofOtherPurpose}
              placeholder='Please specify the purpose of your trip here.'
              onChange={func_setvalueofOtherPurpose}
            />
          </div>
        </div>
        <div className={styles.row} >
          <div className={styles.column20p}>
            Flight Info
          </div>
          <div className={styles.column80p}>
            <AddFlightPanel panelType={PanelType.medium} tempid={tempid} ref={FlightRef} setDirty={setDirty} />
          </div>
        </div>
        <div className={styles.row} >
          <div className={styles.column20p}>
            Airport Pickup?<span className={styles.required}>*</span>
          </div>
          <div className={styles.column80p}>
            <Toggle
              onText="Yes"
              offText="No"
              onChange={_onAirportPickupChange}
              role="checkbox"
              checked={boolAirportPickup}
            />
          </div>
        </div>
        <div className={isCarRentalShow ? styles.row : styles.rownoshow} >
          <div className={styles.column20p}>
            Car Rental?
          </div>
          <div className={styles.column80p}>
            <Toggle
              onText="Yes"
              offText="No"
              onChange={_onCarRentalChange}
              role="checkbox"
              checked={boolCarRental}
            />
          </div>
        </div>
        <div className={styles.row} >
          <div className={styles.column20p} >
            Visitor Point of Contact<span className={styles.required}>*</span>
          </div>
          <div className={styles.column80p} >
            {<EmployeeCombox subprops={passToSubComponent} onValueUpdate={onValueUpdate} ref={childRef} />}
          </div>
        </div>
        <div className={styles.row} >
          <div className={styles.column20p} >
            Mobile Number
          </div>
          <div className={styles.column80p} >
            <div className={styles.subrow} >
              <div className={styles.subrowcolumn100pplus} >
                {<CountryCode subprops={passToSubComponent} onCountryCodeValueUpdate={onCountryCodeValueUpdate} ref={countryCodeRef} />}
                <TextField value={valueofMobile} onChange={func_setvalueofMobile}
                  placeholder="Please include the area code with your mobile number."
                  styles={Common.textStylesofNameBlock}
                />
              </div>
            </div>
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column20p}>
            Meeting Room Needs
        </div>
          <div className={styles.column80p}>
            <TextField multiline rows={6}
              placeholder="Enter the meeting room capacity, VC requirement, the date and the time when you will require the meeting room. This should be booked by you beforehand or by your primary Point of Contact."
              value={valueofRoomNeeds} onChange={func_setvalueofRoomNeeds}
            />
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column20p}>
            Comment
        </div>
          <div className={styles.column80p}>
            <TextField multiline rows={6}
              value={valueofComment} onChange={func_setvalueofComment} />
          </div>
        </div>
        <div className={styles.row}>
          <div className={styles.column20p}>

          </div>
          <div className={styles.column80pRight}>
            <PrimaryButton text="Submit the request" type="Submit" id="btnSubmit" />
            {teachingBubbleVisibleForm && (
              <TeachingBubble
                calloutProps={{ directionalHint: DirectionalHint.rightCenter }}
                target={webControlFocus}
                isWide={true}
                hasCloseButton={true}
                closeButtonAriaLabel="Close"
                onDismiss={toggleTeachingBubbleVisibleForm}
                headline="Attention!"
              >
                {errorMessage}
              </TeachingBubble>
            )}
          </div>
        </div>
      </form>
    </div>
  );

}
