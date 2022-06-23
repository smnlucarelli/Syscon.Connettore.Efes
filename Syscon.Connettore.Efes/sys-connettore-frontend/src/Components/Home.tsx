import { Card } from 'antd';
import React, { Component } from 'react'
import { Container, Row, Col } from 'reactstrap';
import logo from '../logo-sysconanalytics.png'


export default class Home extends Component {
  render() {
    return (
      <>
        <Card>
          <Row>
            {/* <Col className="col-12 text-center">
              <h1>Syscon Analytics</h1>
            </Col> */}
          </Row>
          <Row>
            <Card type='inner' title='Commesse' >

            </Card>
            {/* <Col className="col-12 text-center">
              <img src={logo}/>
            </Col> */}
          </Row>

        </Card>
       </>
    )
  }
}
