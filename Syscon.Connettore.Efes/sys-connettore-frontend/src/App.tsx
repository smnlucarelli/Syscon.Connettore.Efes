import React, { Component } from 'react';
import { Container } from 'reactstrap';
import {BrowserRouter} from 'react-router-dom';
import GeneralLayout from './Components/GeneralLayout';
import Login from './Components/Login';

class App extends Component {
  render() {
    return (
      <BrowserRouter>
        <GeneralLayout></GeneralLayout>
      </BrowserRouter>
    );
  }
}

export default App;