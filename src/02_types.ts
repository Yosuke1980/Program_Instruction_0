// 型定義
interface ProgramResult {
  program: string;
  docId: string;
  url: string;
}

interface MusicItem {
  曲名: string;
  URL: string;
}

interface DaySchedule {
  [key: string]: MusicItem[];
}

interface WeeklySchedule {
  [key: string]: MusicItem[];
  monday: MusicItem[];
  tuesday: MusicItem[];
  wednesday: MusicItem[];
  thursday: MusicItem[];
  friday: MusicItem[];
  saturday: MusicItem[];
  sunday: MusicItem[];
}

interface ProgramData {
  [key: string]: any;
}

interface ContentItem {
  name: string;
  content: string;
}

interface DayData {
  date: string;
  announcements?: ContentItem[];
  music?: ContentItem[];
  guests?: ContentItem[];
  specialFrames?: ContentItem[];
  weeklyContent?: ContentItem[];
}

interface ProgramMetadata {
  name: string;
  programId: string;
  meta: {
    productionCompany?: string;
    genre?: string;
    startDate?: string;
    regularSlots?: string[];
    hosts?: string[];
    notes?: string[];
    publicityFrames?: any[];
    reportFrames?: any[];
    callFrames?: any[];
    specialFrames?: any[];
  };
}

interface CalendarEvent {
  title: string;
  startTime: Date;
  endTime: Date;
  description?: string;
}

interface SpreadsheetRange {
  sheetName: string;
  range: string;
}

interface ProcessingResult {
  success: boolean;
  data?: any;
  error?: string;
  debugLogs?: string[];
}