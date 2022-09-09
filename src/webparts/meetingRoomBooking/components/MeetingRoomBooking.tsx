import * as React from "react";
import styles from "./MeetingRoomBooking.module.scss";
import { IMeetingRoomBookingProps } from "./IMeetingRoomBookingProps";
import { IMeetingRoomBookingState } from "./IMeetingRoomBookingState";
import Modal from "./Modal/Modal";
import * as moment from "moment";
import { graph } from "@pnp/graph";

export default class MeetingRoomBooking extends React.Component<
  IMeetingRoomBookingProps,
  IMeetingRoomBookingState
> {
  constructor(props: IMeetingRoomBookingProps) {
    super(props);
    this.state = {
      events: [],
      eventId: "",
      rooms: [],
      editableCalendar: [],
      isOpened: false,
    };
    this.EditDelete = this.EditDelete.bind(this);
  }
  public async componentDidMount(): Promise<void> {
    this.props.provider
      .getOutlookEvents(this.props.roomsToShow)
      .then((outlookEvents) => {
        this.setState({ events: outlookEvents, eventId: "" });
      });
    this.getRooms();
  }
  public componentDidUpdate(
    prevProps: Readonly<IMeetingRoomBookingProps>,
    prevState: Readonly<IMeetingRoomBookingState>,
    snapshot?: any
  ): void {
    if (prevProps.roomsToShow !== this.props.roomsToShow) {
      this.getRooms();
    }
  }

  private EditDelete(id: string) {
    if (id !== "") {
      this.setState({ eventId: id, isOpened: true });
    } else {
      this.setState({ isOpened: false });
    }
  }
  public getRooms = () => {
    // this.setState({ rooms: rooms });
    let meetingRoomOptions = [];
    if (this.props.roomsToShow !== undefined) {
      for (let i = 0; i < this.props.roomsToShow.length; i++) {
        meetingRoomOptions.push({
          key: this.props.roomsToShow[i],
          text: this.props.roomsToShow[i],
        });
      }

      this.setState({ rooms: meetingRoomOptions });
    }
  };
  public checkIsOpenPanel = (e) => {
    this.setState({ isOpened: e });
  };

  public getEditableCalendar = (allCalendars) => {
    //if the user have the permission to edit a pop will show when the even is clicked, else nothing will happen
    let checkEditableCalendar = allCalendars
      .filter((cal) => cal.canEdit == true)
      .map((res) => res.name);
    this.setState({ editableCalendar: checkEditableCalendar });
  };

  public render(): React.ReactElement<IMeetingRoomBookingProps> {
    return (
      <div className={styles.meetingRoomBooking}>
        <div className={styles.container}>
          <h3>{this.props.title}</h3>

          {this.state.events.length > 0 &&
            this.state.events
              .filter((room) => {
                return this.state.rooms
                  .map((name) => name.text)
                  .includes(room.location.displayName);
              })
              .filter((propertypaneSlectedRoom) => {
                return this.props.roomsToShow.includes(
                  propertypaneSlectedRoom.location.displayName
                );
              })
              .slice(0, 3)
              .map((event) => {
                return (
                  <div
                    className={styles.event_Item}
                    onClick={() =>
                      this.EditDelete(
                        this.state.editableCalendar.includes(
                          event.location.displayName
                        )
                          ? event.id
                          : ""
                      )
                    }
                  >
                    <div className={styles.event_Item_Header}>
                      <div className={styles.event_Item_Header_Title}>
                        <span>{event.subject}</span>
                      </div>
                      <div className={styles.event_Item_Header_Time}>
                        <span>
                          {/*get in this format 10-08-2022 | 10:00 AM - 10:30 AM */}
                          {moment(event.start).format("DD-MM-YYYY")}
                          {event.isAllDay !== true && (
                            <>
                              {" "}
                              | {moment(event.start).format("HH:mm")} -
                              {moment(event.end).format("HH:mm")}
                            </>
                          )}
                        </span>
                      </div>
                    </div>
                    <div className={styles.event_Item_Body}>
                      <div className={styles.event_Item_Body_Location}>
                        <span>{event.location.displayName}</span>
                        {/* <span>{timezoneName}</span> */}
                      </div>
                    </div>
                  </div>
                );
              })}
          {/* <AddEvents
            title={this.props.title}
            provider={this.props.provider}
            context={this.props.context}
          /> */}
          <Modal
            //get rooms from Modal.tsx
            //dropdown rooms
            rooms={this.state.rooms}
            events={this.state.events}
            eventId={this.state.eventId}
            destroyEventId={() => this.setState({ eventId: "" })}
            provider={this.props.provider}
            context={this.props.context}
            isOpened={this.state.isOpened}
            checkIsPanelOpen={this.checkIsOpenPanel}
            getEditableCalendar={this.getEditableCalendar}
            refresh={() => {
              this.componentDidMount();
            }}
            onSave={function (e: any): void {
              throw new Error("Function not implemented.");
            }}
          />
        </div>
      </div>
    );
  }
}
