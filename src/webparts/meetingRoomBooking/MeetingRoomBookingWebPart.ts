import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";

import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "MeetingRoomBookingWebPartStrings";
import MeetingRoomBooking from "./components/MeetingRoomBooking";
import { IMeetingRoomBookingProps } from "./components/IMeetingRoomBookingProps";
import { IServiceProvider } from "./service/IServiceProvider";
import ServiceProvider from "./service/ServiceProvider";
import { sp } from "@pnp/sp";
import { graph } from "@pnp/graph";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls";

let roomsOption = [];
let AllRooms = [];
export interface IMeetingRoomBookingWebPartProps {
  title: string;
  roomsToShow: any[];
}

export default class MeetingRoomBookingWebPart extends BaseClientSideWebPart<IMeetingRoomBookingWebPartProps> {
  private _serviceProvider: IServiceProvider;
  protected onInit(): Promise<void> {
    this._serviceProvider = new ServiceProvider(this.context);

    sp.setup({
      spfxContext: this.context,
    });
    graph.setup({
      spfxContext: this.context,
    });

    return super.onInit();
  }
  public render(): void {
    this.getRooms();
    const element: React.ReactElement<IMeetingRoomBookingProps> =
      React.createElement(MeetingRoomBooking, {
        title: this.properties.title,
        provider: this._serviceProvider,
        context: this.context,
        roomsToShow: this.properties.roomsToShow,
      });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  protected async getRooms(): Promise<any> {
    await graph.me
      .findRooms()
      .get()
      .then((rooms) => {
        roomsOption = [];
        rooms.map((value) => {
          roomsOption.push({
            key: value.name,
            text: value.name,
          });
        });
      });
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Basic Configs",
              groupFields: [
                PropertyPaneTextField("title", {
                  label: "Add a title",
                }),
                PropertyFieldMultiSelect("roomsToShow", {
                  key: "roomsToShow",
                  label: "Select rooms to show",
                  options: roomsOption,
                  selectedKeys: this.properties.roomsToShow,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
