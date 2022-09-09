export interface IModalProps {
  onSave: (e: any) => void;
  refresh: () => void;
  checkIsPanelOpen: (e: boolean) => void;
  getEditableCalendar: (e: any) => void;
  provider: any;
  context: any;
  eventId: string;
  events: any;
  rooms: any;
  destroyEventId: any;
  isOpened: boolean;
}
