import {
  ButtonClickedCallback
} from '../../../models';

export interface IChartTestProps {
  chartData: IChartData;
  onAddListItem? (item: string, percentComplete: number): ButtonClickedCallback;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

export interface IChartData {
  labels: string [];
  datasets: any [];
}