export interface IServiceProvider {
  getMeetingRooms(): Promise<any>;
  getOutlookEvents(roomsToShow): Promise<any>;
  getOutlookCategory(selectedRoom): Promise<any>;
  addMeetingEvent(calendarID, event: any): Promise<any>;
  updateMeetingEvent(id: string, event: any): Promise<any>;
  getbyId(id: string): Promise<any>;
}
