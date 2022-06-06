import * as React from 'react';
import Dropdown from 'react-dropdown';
import 'react-dropdown/style.css';
import styles from './ChartTest.module.scss';

import { IChartTestProps } from './IChartTestProps';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
} from 'chart.js';
import { Bar, Pie } from 'react-chartjs-2';

import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'ChartTestWebPartStrings';

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
);

export const optionsMenu = [
  'Bar Vertical', 'Bar Horizontal'
];

interface IChartTestState {
  selected: string;
  item: string;
  percentComplete: number;
}

const optionDefault = optionsMenu[0];

export default class ChartTest extends React.Component<IChartTestProps, IChartTestState> {
  constructor (props) {
    super(props);
    this.state = {
      selected: 'Bar Vertical',
      item: "",
      percentComplete: 0
    };
    this._onSelect = this._onSelect.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
    this.handleChangeItem = this.handleChangeItem.bind(this);
    this.handleChangePercentComplete = this.handleChangePercentComplete.bind(this);
  }

  private _onSelect (option) {
    this.setState({selected: option.label});
  }

  private handleSubmit(event) {
    this.props.onAddListItem(this.state.item, this.state.percentComplete);
    event.preventDefault();
  }

  private handleChangeItem(event) {
    this.setState({item: event.target.value});
  }

  private handleChangePercentComplete(event) {
    this.setState({percentComplete: event.target.value});
  }

  private chart() {
    
    switch (this.state.selected){
      case "Bar Vertical":
        const options = {
          responsive: true,
          plugins: {
            legend: {
              position: 'top' as const,
            },
            title: {
              display: true,
              text: 'Vertical Bar Char',
            },
          },
        };
        return (<Bar options={options} data={this.props.chartData} />);

      case "Bar Horizontal":
        const optionsBar = {
          indexAxis: 'y' as const,
          elements: {
            bar: {
              borderWidth: 2,
            },
          },
          responsive: true,
          plugins: {
            legend: {
              position: 'right' as const,
            },
            title: {
              display: true,
              text: 'Horizontal Bar Chart',
            },
          },
        };
        return(<Bar options={optionsBar} data={this.props.chartData} />);
    }
  }

  public render(): React.ReactElement<IChartTestProps> {

    const {
      chartData,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <div>
        <h3>Tipo de Grafico</h3>
        <Dropdown 
          options={optionsMenu} 
          onChange={this._onSelect} 
          value={optionDefault} 
          placeholder="Select an option" 
        />

        <br></br>
        <div>{ this.chart() }</div>
        <br></br>

        <h3>Agregar Items</h3>
        <form onSubmit={this.handleSubmit}>
          <label>
            Work Item
            <input type="text" value={this.state.item} onChange={this.handleChangeItem} />
          </label>
          <label>
            Percent Complete
            <input type="text" value={this.state.percentComplete} onChange={this.handleChangePercentComplete} />
          </label>
          <input type="submit" value="Submit" />
        </form>

      </div>
    );
  }
}

