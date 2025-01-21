import { DatePicker, DatePickerProps } from '@mui/x-date-pickers/DatePicker';
import { TimePicker, TimePickerProps } from '@mui/x-date-pickers/TimePicker';
import { styled } from '@mui/material/styles';
import { Moment } from 'moment';

interface CustomPickerProps {
  error?: boolean;
}
export interface SchemaField {
  name: string; // The field name
  label: string; // The display label for the field
  type: 'string' | 'number' | 'date' | 'action'; // The field type
  width?: number;
  component?: JSX.Element;
  editable?: boolean;
  action?: Function
}

export interface Schema {
  fields: SchemaField[];
}

export interface BaseEntity {
  id: number;
  uid?: string;
  idView?: string;
}

export interface Task extends BaseEntity {
  name: string;
  company: string;
  designation: string;
  department: string;
  subject: string;
  startDate: string;
  endDate: string;
  importance: string;
  description: string;
  attachedId?: string;
  ids?: string[];
  forInfo?: string;
  forInfoIds?: string[];
}

export interface Employee extends BaseEntity {
  name: string;
  company: string;
  designation: string;
}

export interface MeetingContent extends BaseEntity {
  description: string;
  name: string;
  dueDate: string;
  status: string;
  ids?: any[]
}

export type SchemaType = 'Task' | 'Employee' | 'MeetingContent';

export type Entity = { name?: string; company?: string; Title?: string };

// Higher-order function to create styled components
const createStyledPicker = (PickerComponent: any) => styled(PickerComponent)<CustomPickerProps & DatePickerProps<unknown> & TimePickerProps<unknown>>(({ theme, error }) => ({
  '& .MuiOutlinedInput-root': {
    '& fieldset': {
      borderColor: error ? theme.palette.error.main : undefined,
    },
    '&:hover fieldset': {
      borderColor: error ? theme.palette.error.main : undefined,
    },
    '&.Mui-focused fieldset': {
      borderColor: error ? theme.palette.error.main : undefined,
    },
  },
  '& .MuiInputLabel-root': {
    color: error ? theme.palette.error.main : undefined,
  },
}));

export const StyledDatePicker = createStyledPicker(DatePicker);
export const StyledTimePicker = createStyledPicker(TimePicker);
