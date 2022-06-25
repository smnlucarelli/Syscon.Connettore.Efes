import React, { Component } from 'react';
import {BrowserRouter} from 'react-router-dom';
import Container from './Components/Container';
import Login from './Components/login/Login';

class App extends Component {
  render() {
    return (
      <BrowserRouter>
        <Container></Container>
      </BrowserRouter>
    );
  }
}

export default App;