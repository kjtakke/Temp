//Arrays
let tableAray = [
  ["Name", "Age", "km Traveled"],
  ["John", 24, 18000],
  ["Ryan", 30, 30000],
  ["George", 28, 41000],
  ["Kym", 31, 28000]
];

//React Variables
let myVariables = {
  heading: "Income Table",
  IncomeTable: tableAray //arrayToHTMLTable(tableAray)
}

let myTags = {
  heading: "welcome",
  IncomeTable: "IncomeTable"
}

//functions

const arrayToTable = (arayData) => {

  //Header
  const headerItems = []
  for (var i = 0; i < 1; i++) {
    const headerItem = []
    for (var j = 0; j < arayData[i].length; j++) {
      headerItem.push(<th>{arayData[i][j]}</th>)
    }
      headerItems.push(<tr>{headerItem}</tr>)
  }

  //Body
  const bodyItems = []
  for (var i = 1; i < arayData.length; i++) {
    const bodyItem = []
    for (var j = 0; j < arayData[i].length; j++) {
      bodyItem.push(<td>{arayData[i][j]}</td>)
    }
      bodyItems.push(<tr>{bodyItem}</tr>)
  }
  return (

    <table class='table table-hover'>
      <thead>
        {headerItems}
      </thead>
      <tbody>
        {bodyItems}
      </tbody>
    </table>
  )
}

//Classes
class Base extends React.Component {
  render() {
    return (
      <div>
        <div id={this.props.heading}></div>
        <div id={this.props.IncomeTable}></div>
      </div>
    );
  }
}

class Heading extends React.Component {
  render() {
    return (
        <h1>{this.props.heading_txt}</h1>
    );
  }
}

class IncomeTable extends React.Component {
  render() {

    return (
        <div>
          {arrayToTable(this.props.IncomeTable)}
        </div>
    );
  }
}

//Render to DOM
ReactDOM.render(
  <Base
    heading={myTags.heading}
    IncomeTable={myTags.IncomeTable}
  />,
  document.getElementById('root')
)
  //Render to Sub Elements
  ReactDOM.render(
    <Heading heading_txt={myVariables.heading}/>,
    document.getElementById(myTags.heading)
  )

  ReactDOM.render(
    <IncomeTable
      IncomeTable={myVariables.IncomeTable}
    />,
    document.getElementById(myTags.IncomeTable)
  )
