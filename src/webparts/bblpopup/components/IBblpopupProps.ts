import { IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';

export interface IBblpopupProps {
  url:string;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  eventStartDate:  IDateTimeFieldValue;
  eventEndDate: IDateTimeFieldValue;
  webpartid :string;

}
