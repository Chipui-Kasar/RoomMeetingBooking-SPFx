import { sp, Web } from "@pnp/sp/presets/all";
import { IServiceProvider } from "./IServiceProvider";
import { MSGraphClient } from "@microsoft/sp-http";
import { BaseWebPartContext } from "@microsoft/sp-webpart-base";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/calendars";
import "@pnp/graph/outlook";
import { Calendars } from "@pnp/graph/calendars";
import * as moment from "moment-timezone";
export default class ServiceProvider implements IServiceProvider {
  private _webPartContext: BaseWebPartContext;
  private _webAbsoluteUrl: string;

  constructor(_context: BaseWebPartContext) {
    this._webPartContext = _context;
    this._webAbsoluteUrl = _context.pageContext.web.absoluteUrl;
  }
  public async getOutlookEvents(roomsToShow): Promise<any> {
    let offset = new Date().getTimezoneOffset();
    let allEvents = [];
    // //get all events from outlook calendar
    // let events = await graph.me.events.get();
    if (roomsToShow !== undefined) {
      const myCalendars = await graph.me.calendars();
      let newCalendar = myCalendars.filter((room) =>
        roomsToShow.includes(room.name)
      );

      for (let i = 0; i < newCalendar.length; i++) {
        let events = await graph.me.calendars
          .getById(newCalendar[i].id)
          .events.filter("isCancelled eq false")();
        events.map((event) => {
          let Aevent = {
            attendees: event.attendees,
            body: event.body,
            bodyPreview: event.bodyPreview,
            categories: event.categories,
            changeKey: event.changeKey,
            createdDateTime: event.createdDateTime,
            end: moment
              .utc(event.end.dateTime)
              .utcOffset(-offset)
              .format("YYYY-MM-DDTHH:mm:ss"),
            id: event.id,
            isAllDay: event.isAllDay,
            isOnlineMeeting: event.isOnlineMeeting,
            isOrganizer: event.isOrganizer,
            isReminderOn: event.isReminderOn,
            location: event.location,
            locations: event.locations,
            onlineMeeting: event.onlineMeeting,
            onlineMeetingProvider: event.onlineMeetingProvider,
            onlineMeetingUrl: event.onlineMeetingUrl,
            organizer: event.organizer,
            originalStartTimeZone: event.originalStartTimeZone,
            recurrence: event.recurrence,
            start: moment
              .utc(event.start.dateTime)
              .utcOffset(-offset)
              .format("YYYY-MM-DDTHH:mm:ss"),
            subject: event.subject,
          };
          allEvents.push(Aevent);
        });
      }
    }

    // events.map((event) => {
    //   let Aevent = {
    //     attendees: event.attendees,
    //     body: event.body,
    //     bodyPreview: event.bodyPreview,
    //     categories: event.categories,
    //     changeKey: event.changeKey,
    //     createdDateTime: event.createdDateTime,
    //     end: moment
    //       .utc(event.end.dateTime)
    //       .utcOffset(-offset)
    //       .format("YYYY-MM-DDTHH:mm:ss"),
    //     id: event.id,
    //     isAllDay: event.isAllDay,
    //     isOnlineMeeting: event.isOnlineMeeting,
    //     isOrganizer: event.isOrganizer,
    //     isReminderOn: event.isReminderOn,
    //     location: event.location,
    //     locations: event.locations,
    //     onlineMeeting: event.onlineMeeting,
    //     onlineMeetingProvider: event.onlineMeetingProvider,
    //     onlineMeetingUrl: event.onlineMeetingUrl,
    //     organizer: event.organizer,
    //     originalStartTimeZone: event.originalStartTimeZone,
    //     recurrence: event.recurrence,
    //     start: moment
    //       .utc(event.start.dateTime)
    //       .utcOffset(-offset)
    //       .format("YYYY-MM-DDTHH:mm:ss"),
    //     subject: event.subject,
    //   };
    //   allEvents.push(Aevent);
    // });
    return allEvents;
  }
  public async getOutlookCategory(selectedRoom): Promise<any> {
    return await graph.me.outlook.masterCategories.get();
  }
  public async getMeetingRooms(): Promise<any> {
    return await graph.me.findRooms().get();
    // return group;
  }
  public addMeetingEvent(calendarID, event: any): Promise<any> {
    //adding in specific calendar
    return graph.me.calendars.getById(calendarID).events.add(event);
  }
  public updateMeetingEvent(id, event: any): Promise<any> {
    return graph.me.events.getById(id).update(event);
  }
  public deleteMeetingEvent(calendarId, id: string): Promise<any> {
    return graph.me.events.getById(id).delete();
  }

  public async getbyId(id: string): Promise<any> {
    return await graph.me.events.getById(id).get();
  }
}
