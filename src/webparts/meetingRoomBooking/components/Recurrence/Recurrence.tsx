import * as moment from "moment";
import { Dropdown, IDropdownOption, TextField } from "office-ui-fabric-react";
import * as React from "react";
import { IRecurrenceProps } from "./IRecurrenceProps";
let daysName = [];
let days = moment.weekdays();
days.forEach((day) => {
  daysName.push({ key: day, text: day });
});

//get month names using js
let months = moment.months();
let monthNames = [];
months.forEach((month) => {
  monthNames.push({ key: month, text: month });
});
let repeatNames = [
  { key: "daily", text: "Daily" },
  { key: "weekly", text: "Weekly" },
  { key: "absoluteMonthly", text: "Monthly" },
];
let selectedDays = [];

function Recurrence(props: IRecurrenceProps) {
  const [recurrenceOn, setRecurrenceOn] = React.useState({
    key: "",
    text: "",
  });
  const [recurrenceType, setRecurrenceType] = React.useState("");
  React.useEffect(() => {}, [props.week]);

  const onSelectWeek = (
    event: React.FormEvent<HTMLDivElement>,
    item?: IDropdownOption
  ) => {
    selectedDays.push(item);
    if (item.selected == false) {
      selectedDays = selectedDays.filter((r) => r.key !== item.key);
    }
    let final = selectedDays.filter((e) => e.selected == true);
    //convert to array of object to array
    let finalArray = final.map((e) => e.key);

    if (final.length > 0) {
      props.selectedWeek(finalArray);
    } else {
      props.selectedWeek([]);
    }
  };

  return (
    <>
      {props.recurrence == true && (
        <Dropdown
          label="Repeat"
          selectedKey={recurrenceOn.key}
          onChange={(
            ev: React.FormEvent<HTMLDivElement>,
            item: IDropdownOption | any
          ): void => {
            setRecurrenceOn(item);
            props.RecurrenceOn(item);
            setRecurrenceType(item.key);
            props.RecurrenceType(item.key);
          }}
          defaultValue={recurrenceOn.key}
          // style={{ width: "60px" }}
          options={repeatNames}
          errorMessage={recurrenceOn.key == "" ? "Required" : ""}
        />
      )}
      {recurrenceType !== "daily" && recurrenceType !== "" && (
        <>
          {recurrenceType == "weekly" && (
            <Dropdown
              label="Every"
              // selectedKey={selectedWeek}
              onChange={onSelectWeek}
              // style={{ width: "60px" }}
              options={daysName}
              multiSelect
              errorMessage={props.week.length == 0 ? "Required" : ""}
            />
          )}
          {recurrenceType == "absoluteMonthly" && (
            <TextField
              name="Monthly"
              value={`Every ${moment(props.startDate).format(
                "Do"
              )} of the month`}
              onGetErrorMessage={(value: string) => {
                if (value.length == 0) {
                  return "Select the Start Date";
                }
              }}
              readOnly
            />
          )}
        </>
      )}
    </>
  );
}

export default Recurrence;
