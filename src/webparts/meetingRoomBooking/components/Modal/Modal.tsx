import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import * as moment from "moment";
import * as moment from "moment-timezone";
import {
  DatePicker,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Panel,
  Spinner,
  TextField,
  Toggle,
} from "office-ui-fabric-react";

import * as React from "react";
import { toLocaleShortDateString } from "../../utils/dateUtils";
import { IMeetingRoomBookingProps } from "../IMeetingRoomBookingProps";
import Recurrence from "../Recurrence/Recurrence";
import { IModalProps } from "./IModalProps";
import styles from "./Modal.module.scss";
import "./index.css";
import { graph } from "@pnp/graph";

let meetingRoomOptions: IDropdownOption[] = [];
let categoryOptions: IDropdownOption[] = [];

let selectedAttendees = [];

const hours = [];
for (let hour = 0; hour < 24; hour++) {
  hours.push({
    key: moment({ hour }).format("HH:mm"),
    text: moment({ hour }).format("HH:mm"),
  });
  hours.push({
    key: moment({
      hour,
      minute: 30,
    }).format("HH:mm"),
    text: moment({
      hour,
      minute: 30,
    }).format("HH:mm"),
  });
}

let eventId = "";
let attendeesUsers = [];
//setting the time from now nearest to digit divisible by 30
const start = moment();
const remainder = 30 - (start.minute() % 30);
const sdateTime = moment(start).add(remainder, "minutes").format("HH:mm");

function Modal(props: IModalProps) {
  const [accessDenied, setAccessDenied] = React.useState("");
  const [globalErrorMessage, setGlobalErrorMesssage] = React.useState("");
  const [rooms, setRooms] = React.useState([]);
  const [title, setTitle] = React.useState("");

  const [recurrence, setRecurrence] = React.useState(false);
  const [selectedWeek, setSelectedWeek] = React.useState([]);
  const [recurrenceOn, setRecurrenceOn] = React.useState({
    key: "",
    text: "",
  });
  const [recurrenceType, setRecurrenceType] = React.useState("");
  // const [selectedCategory, setSelectedCategory] = React.useState("");
  const [allDayEvent, setAllDayEvent] = React.useState(false);
  const [startDate, setStartDate] = React.useState(new Date());
  //set null
  const [endDate, setEndDate] = React.useState(null);
  const [startHour, setStartHour] = React.useState({
    key: sdateTime,
    text: sdateTime,
  });

  const [endHour, setEndHour] = React.useState({
    key: "00:00",
    text: "00:00",
  });

  const [selectedRoom, setSelectedRoom] = React.useState("");
  const [calendarSavedRoom, setCalendarSavedRoom] = React.useState("");
  const [attendees, setAttendees] = React.useState([]);
  const [selectedAtendees, setSelectedAttendees] = React.useState([]);
  const [description, setDescription] = React.useState("");
  const [statusMessage, setStatusMessage] = React.useState("");
  const [saveStatus, setSaveStatus] = React.useState(false);
  const [deleteStatus, setDeleteStatus] = React.useState(false);
  const [calendarId, setCalendarId] = React.useState("");
  // const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] =
  //   useBoolean(false);
  const [disableRoom, setDisableRoom] = React.useState(false);
  const [isOpened, setIsOpened] = React.useState(false);

  eventId = props.eventId;
  //Date Formatting
  let startTime = `${startHour.key}`;
  let startDateTime = `${moment(startDate).format("YYYY-MM-DD")}T${
    !allDayEvent ? startTime : "00:00:00"
  }`;
  let endTime = `${endHour.key}`;
  let endDateTime = `${moment(allDayEvent == true ? startDate : endDate)
    .add(allDayEvent == true ? 1 : 0, "day")
    .format("YYYY-MM-DD")}T${!allDayEvent ? endTime : "00:00:00"}`;

  React.useEffect(() => {
    if (
      title !== "" &&
      attendees.length > 0 &&
      !endDateTime.includes("Invalid dateT00:00") &&
      selectedRoom !== ""
    ) {
      setGlobalErrorMesssage("");
    }
    //Checking whether the time slot is available or not
    props.events
      .filter((item) => {
        if (
          item.location.displayName === selectedRoom &&
          (moment(startDateTime).isBetween(
            moment(item.start).subtract(1, "minute"),
            moment(item.end).add(1, "minute")
          ) ||
            moment(endDateTime).isBetween(
              moment(item.start).subtract(1, "minute"),
              moment(item.end).add(1, "minute")
            )) &&
          // moment(item.start).isBetween(startDateTime, endDateTime) ||
          moment(item.start).format("YYYY-MM-DDTHH:mm") == startDateTime &&
          moment(item.end).format("YYYY-MM-DDTHH:mm") == endDateTime
        ) {
          return item;
        } else {
          setStatusMessage("");
          return;
        }
      })
      .map((item) => {
        if (eventId == "") {
          setStatusMessage("Room is already booked at the selected time");
        } else {
          setStatusMessage("");
        }
      });
  }, [props.events, selectedRoom, statusMessage, startDateTime, endDateTime]);
  //using different hooks for different states
  React.useEffect(() => {
    let isCancelled = false;
    //this will run only if we click on specific item in the list
    if (eventId !== "" && !isCancelled) {
      let offset = new Date().getTimezoneOffset();
      props.provider
        .getbyId(eventId)
        .then((event) => {
          selectedAttendees = [];
          event.attendees.forEach((value) => {
            selectedAttendees.push({
              emailAddress: {
                address: value.emailAddress.address,
                name: value.emailAddress.name,
              },
              type: "required",
            });
          });

          //getting the attendees
          attendeesUsers.push(
            event.attendees.map((attendee) => attendee.emailAddress.address)
          );
          setTitle(event.subject);
          setAllDayEvent(event.isAllDay);
          setStartDate(new Date(event.start.dateTime));
          setEndDate(new Date(event.end.dateTime));
          // setSelectedCategory(event.categories[0]);
          setStartHour({
            key: moment
              .utc(event.start.dateTime)
              .utcOffset(-offset)
              .format("HH:mm"),
            text: moment
              .utc(event.start.dateTime)
              .utcOffset(-offset)
              .utcOffset(-offset)
              .format("HH:mm"),
          });

          setEndHour({
            key: moment
              .utc(event.end.dateTime)
              .utcOffset(-offset)
              .format("HH:mm"),
            text: moment
              .utc(event.end.dateTime)
              .utcOffset(-offset)
              .format("HH:mm"),
          });

          setSelectedRoom(event.location.displayName);
          setCalendarSavedRoom(event.location.displayName);
          // setAttendees(attendeesUsers[0]);
          setAttendees(selectedAttendees);
          setSelectedAttendees(attendeesUsers[0]);
          setDescription(event.bodyPreview);
        })
        .catch((error) => {
          console.log(error);
        });
    } else {
      // setRooms(props.rooms);
      setDisableRoom(true);
    }
    //checking if the user has room calendar
    checkCalendarRoom();
    return () => {
      isCancelled = true;
    };
  }, [eventId, props.rooms]);
  React.useEffect(() => {
    // getOutlookCategory();
  }, []);
  React.useEffect(() => {
    setIsOpened(props.isOpened);
  }, [props.isOpened]);

  // const getOutlookCategory = () => {
  //   props.provider.getOutlookCategory(selectedRoom).then((categories) => {
  //     categoryOptions = [];
  //     categories.map((value) => {
  //       categoryOptions.push({
  //         key: value.displayName,
  //         text: value.displayName,
  //       });
  //     });
  //     setCategory(categoryOptions);
  //   });
  // };

  // When the user clicks anywhere outside of the modal, close it

  const getAttendees = async (attendees: any[]) => {
    selectedAttendees = [];
    attendees.forEach((value) => {
      selectedAttendees.push({
        emailAddress: {
          address: value.id.split("|")[2],
          name: value.text,
        },
        type: "required",
      });
    });
    setSelectedAttendees(selectedAttendees);
    setAttendees(selectedAttendees);
  };

  const RoomChange = async (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number
  ): Promise<void> => {
    setSelectedRoom(option.text);

    const myCalendars = await graph.me.calendars();
    let filteredCalendar = myCalendars.filter(
      (room) => room.name === option.text
    );
    setCalendarId(filteredCalendar[0].id);
  };
  const checkCalendarRoom = async () => {
    const myCalendars = await graph.me.calendars();
    // let filteredCalendar = myCalendars.filter((room) =>
    //   props.rooms.map((res) => res.key).includes(room.name)
    // );
    props.getEditableCalendar(myCalendars);
    //filtering only rooms calendar
    let filteredCalendar = props.rooms.filter((room) =>
      myCalendars.map((rooms) => rooms.name).includes(room.key)
    );

    if (filteredCalendar.length == 0) {
      setDisableRoom(true);
    } else {
      //filtering out only the editable room calendars
      let EditableCalendar = props.rooms.filter((room) =>
        myCalendars
          .filter((edit) => edit.canEdit == true)
          .map((rooms) => rooms.name)
          .includes(room.key)
      );
      setRooms(EditableCalendar);
      setDisableRoom(false);
    }
  };
  // const categoryChange = (
  //   event: React.FormEvent<HTMLDivElement>,
  //   option?: IDropdownOption,
  //   index?: number
  // ): void => {
  //   setSelectedCategory(option.text);
  // };
  const onSave = async () => {
    if (startDateTime >= endDateTime) {
      alert("Start Date & time must be before end Date & time");
      return;
    }
    //get browser timezone like India Standard Time or Universal Time
    var current_date = new Date();
    var n = current_date.toString();
    var arr = n.split("(");
    var result = arr[arr.length - 1];
    var timezone = result.replace(")", "");
    //generate ms teams meeting link
    // const w = await (await sp.getTenantAppCatalogWeb()).get();
    // const user = await graph.me();
    // //generate random ms teams meeting id
    // const meetingId = Math.random().toString(36).substring(2, 15);
    // var meetingLink = `https://teams.microsoft.com/l/meetup-join/${w.Id}/${meetingId}?context=%7b%22Tid%22%3a%22${w.Id}%22%2c%22Oid%22%3a%22${user.id}%22%7d`;

    let items = {
      subject: title,
      body: {
        contentType: "HTML",
        content: description,
      },
      recurrence:
        recurrence == true
          ? {
              pattern: {
                dayOfMonth:
                  recurrenceType == "absoluteMonthly"
                    ? moment(startDate).format("D")
                    : 0,
                daysOfWeek: selectedWeek,
                firstDayOfWeek: "sunday",
                index: "first",
                interval: 1,
                month: 0,
                type: recurrenceType,
              },
              range: {
                endDate: "0001-01-01",
                numberOfOccurrences: 0,
                // recurrenceTimeZone: "India Standard Time",
                startDate: moment(startDate).format("YYYY-MM-DD"),
                type: "noEnd",
              },
            }
          : null,
      isAllDay: allDayEvent ? true : false,
      isOnlineMeeting: true,

      start: {
        dateTime: startDateTime,
        timeZone: timezone,
      },
      end: {
        dateTime: endDateTime,
        timeZone: timezone,
      },
      location: {
        displayName: selectedRoom,
      },
      // categories: [selectedCategory],
      attendees: attendees.length > 0 ? attendees : selectedAtendees,
      id: Date.now().toString(36) + Math.random().toString(36).substr(2),
    };

    if (
      title !== "" &&
      attendees.length > 0 &&
      !endDateTime.includes("Invalid") &&
      selectedRoom !== ""
    ) {
      setSaveStatus(true);
      if (eventId && calendarSavedRoom === selectedRoom) {
        props.provider
          .updateMeetingEvent(eventId, items)
          .then((res) => {
            // window.location.reload();
            props.refresh();
            setSaveStatus(false);
            dismissPanel();
          })
          .catch((error) => {
            console.log(error);
            let split = error.toString().split("::>")[1];
            if (split.includes("Access is denied")) {
              setAccessDenied("Access is Denied");
            }
          });
      } else if (eventId && calendarSavedRoom !== selectedRoom) {
        props.provider
          .addMeetingEvent(calendarId, items)
          .then((res) => {
            onDelete();
            props.refresh();

            setSaveStatus(false);
            dismissPanel();
          })
          .catch((error) => {
            console.log(error);
            let split = error.toString().split("::>")[1];
            if (split.includes("Access is denied")) {
              setAccessDenied("Access is Denied");
            }
          });
      } else {
        props.provider
          .addMeetingEvent(calendarId, items)
          .then((res) => {
            props.refresh();

            setSaveStatus(false);
            dismissPanel();
          })
          .catch((error) => {
            console.log(error);
            let split = error.toString().split("::>")[1];
            if (split.includes("Access is denied")) {
              setAccessDenied("Access is Denied");
            }
          });
      }
    } else {
      setGlobalErrorMesssage("Please fill all the required fields");
    }
  };
  const onDelete = () => {
    setDeleteStatus(true);
    props.provider
      .deleteMeetingEvent(calendarId, eventId)
      .then((res) => {
        dismissPanel();
        props.refresh();
        setDeleteStatus(false);
      })
      .catch((error) => {
        console.log(error);
        let split = error.toString().split("::>")[1];
        if (split.includes("Access is denied")) {
          setAccessDenied("Access is Denied");
        }
      });
  };

  const openPanel = () => {
    setIsOpened(true);
  };
  const dismissPanel = () => {
    setAccessDenied("");
    setIsOpened(false);
    props.checkIsPanelOpen(false);
    props.destroyEventId();
    // document.getElementById("myModal").style.display = "block";
    setAllDayEvent(false);
    setTitle("");
    setStartDate(new Date());
    // setSelectedCategory("");
    setRecurrence(false);
    setSelectedAttendees([]);
    setSelectedWeek([]);
    setRecurrenceOn({ key: "", text: "" });
    setRecurrenceType("");
    setEndDate(null);
    setStartHour({ key: "00", text: "00" });

    setEndHour({ key: "00", text: "00" });

    setSelectedRoom("");
    setAttendees([]);
    setDescription("");
    setSaveStatus(false);
    setDeleteStatus(false);
  };

  return (
    <div className={styles.addEventContainer}>
      <button onClick={openPanel} className={styles.openModalBtn}>
        Book New Meeting
      </button>
      <Panel
        isOpen={isOpened}
        onDismiss={dismissPanel}
        headerText="Book a Meeting"
        closeButtonAriaLabel="Close"
        isFooterAtBottom={true}
        // isLightDismiss={true}
      >
        {globalErrorMessage && (
          <div style={{ color: "#a80000", fontSize: "12px" }}>
            {globalErrorMessage}
          </div>
        )}
        <TextField
          label="Event Name"
          name="Title"
          value={title}
          onChange={(
            event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
            newValue?: string
          ) => setTitle(newValue)}
          required
          onGetErrorMessage={(value: string) => {
            if (value.length == 0) {
              return "Title is required";
            }
          }}
        />
        <Dropdown
          label="Event Location"
          placeholder="Select an option"
          options={rooms}
          disabled={disableRoom}
          selectedKey={selectedRoom}
          // multiSelect
          onChange={RoomChange}
          required
        />
        {disableRoom == true && (
          <div style={{ color: "#a80000", fontSize: "12px" }}>
            Access is required for the meeting rooms, Please contact your admin.
          </div>
        )}
        {statusMessage !== "" && (
          <div style={{ color: "#a80000", fontSize: "12px" }}>
            {statusMessage}
          </div>
        )}
        <PeoplePicker
          titleText="Attendees"
          // webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
          placeholder="Select Attendees"
          principalTypes={[PrincipalType.User]}
          context={props.context}
          personSelectionLimit={20}
          groupName={""}
          showtooltip={true}
          defaultSelectedUsers={selectedAtendees}
          onChange={getAttendees}
          required
          errorMessage={
            attendees.length == 0 ? "Please select at least one attendee" : ""
          }
          // onGetErrorMessage={(value: any[]) => {
          //   if (value.length == 0) {
          //     return "Attendees is required";
          //   }
          // }}
        />
        {/* <Dropdown
          label="Event Type"
          selectedKey={selectedCategory}
          onChange={categoryChange}
          defaultValue={selectedCategory}
          // style={{ width: "60px" }}
          options={category}
        /> */}
        <TextField
          label="Event Description"
          multiline
          name="Description"
          value={description}
          onChange={(
            event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
            newValue?: string
          ) => setDescription(newValue)}
        />
        <div className={styles.comboItem}>
          <Toggle
            label="Recurrence"
            defaultChecked={recurrence}
            onText="Yes"
            offText="No"
            onChange={(event, checked) => {
              setRecurrence(checked);
            }}
          />
          <Toggle
            label="All Day"
            defaultChecked={allDayEvent}
            onText="Yes"
            offText="No"
            onChange={(event, checked) => {
              setAllDayEvent(checked);
            }}
          />
        </div>
        <Recurrence
          recurrence={recurrence}
          startDate={startDate}
          selectedWeek={setSelectedWeek}
          week={selectedWeek}
          RecurrenceOn={setRecurrenceOn}
          RecurrenceType={setRecurrenceType}
        />
        <div className={styles.comboItem}>
          <DatePicker
            label="Start Date & Time"
            isRequired={true}
            // strings={DayPickerStrings}
            placeholder="Select Date"
            allowTextInput={true}
            value={startDate}
            formatDate={toLocaleShortDateString}
            onSelectDate={(date: Date) => setStartDate(date)}
            style={{ marginRight: "5px" }}
            minDate={new Date()}
          />
          {!allDayEvent && (
            <>
              <Dropdown
                required
                selectedKey={startHour.key}
                onChange={(
                  ev: React.FormEvent<HTMLDivElement>,
                  item: IDropdownOption | any
                ): void => {
                  setStartHour(item);
                }}
                defaultValue={startHour.text}
                style={{ width: "80px" }}
                options={hours}
              />
            </>
          )}
        </div>
        {!allDayEvent && (
          <div className={styles.comboItem}>
            <DatePicker
              label="End Date & Time"
              isRequired={true}
              // strings={DayPickerStrings}
              placeholder="Select Date"
              allowTextInput={true}
              value={endDate}
              formatDate={toLocaleShortDateString}
              onSelectDate={(date: Date) => setEndDate(date)}
              style={{ marginRight: "5px" }}
              minDate={startDate}
              onError={(error: any) => console.log(error)}
            />

            <Dropdown
              required
              selectedKey={endHour.key}
              onChange={(
                ev: React.FormEvent<HTMLDivElement>,
                item: IDropdownOption | any
              ): void => {
                setEndHour(item);
              }}
              defaultValue={endHour.key}
              style={{ width: "80px" }}
              options={hours}
            />
          </div>
        )}

        <div
          style={{
            display: "flex",
            justifyContent: "space-around",
            paddingTop: "20px",
          }}
        >
          <DefaultButton
            type="submit"
            text={"Submit"}
            disabled={disableRoom}
            onClick={onSave}
            className={styles.saveItem}
            style={{
              cursor:
                saveStatus == true || disableRoom == true
                  ? "not-allowed"
                  : "cursor",
              background:
                saveStatus == true || disableRoom == true ? "#D3D3D3" : "",
            }}
          />
          <DefaultButton
            type="submit"
            text={"Cancel"}
            onClick={dismissPanel}
            className={styles.saveItem}
          />
          {eventId && (
            <DefaultButton
              type="submit"
              text={"Delete"}
              onClick={onDelete}
              disabled={disableRoom}
              className={styles.deleteItem}
              style={{
                cursor:
                  deleteStatus == true || saveStatus == true
                    ? "not-allowed"
                    : "cursor",
              }}
            />
          )}
        </div>
        {(saveStatus == true || deleteStatus == true) && (
          <div className={styles.Spinner}>
            {accessDenied == "" ? (
              <Spinner
                label={
                  saveStatus == true
                    ? "Saving Please Wait..."
                    : "Deleting Please Wait..."
                }
                ariaLive="assertive"
                labelPosition="right"
              />
            ) : (
              <div style={{ display: "flex", flexDirection: "column" }}>
                <div className={styles.accessDeniedMsg}>{accessDenied}</div>

                <DefaultButton
                  type="submit"
                  text={"Close"}
                  onClick={dismissPanel}
                  className={styles.accessDeniedBtn}
                />
              </div>
            )}
          </div>
        )}
      </Panel>
    </div>
  );
}

export default Modal;
