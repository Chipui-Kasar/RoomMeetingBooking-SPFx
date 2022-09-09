import { IServiceProvider } from "../service/IServiceProvider";

export interface IMeetingRoomBookingProps {
  title: string;
  provider: IServiceProvider;
  context: any;
  roomsToShow: any[];
}
