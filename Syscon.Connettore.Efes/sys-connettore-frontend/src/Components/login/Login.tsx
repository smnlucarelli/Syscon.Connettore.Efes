import { LockOutlined, UserOutlined } from '@ant-design/icons'
import { Button, Col, Form, Input, Modal, Row } from 'antd'
import Title from 'antd/lib/typography/Title'
import axios from 'axios'
import React, { Component } from 'react'
import logo3 from '../logo-sysconanalytics.png'
import Container from '../Container'


export default class Login extends Component<any, any> {
    
    constructor(props) {
        super(props);

        this.state = { 
            login: false,
            user: [],
            userlogin: '',
            passlogin: ''
        }


      }


    onChangeUser = (e) => {
        this.setState({ userlogin: e.target.value })
    }

    onChangePassword = (e) => {
        this.setState({ passlogin: e.target.value })
    }

    onLogin = (username: string, password: string) => {
    
        const rlogin = axios.get('https://localhost:44369/api/v1/Login?username=' + username + '&password=' + password, { headers: {'Accept': 'application/json','Content-Type': 'application/json'}})
                            .then((response) => {
                                const login = response.data;
                                const user = response.data.item2;

                                this.setState({ login, user })
                            })

    }

    render() {

        let { userlogin, passlogin, } = this.state;


        let authRedirect;
        if (this.state.login.item1 == true) {
            return ( <Container></Container>)
        }

        return(
            <>
                <Modal visible closable={false} footer={null}>
                    <Row>
                        <Col className="col-12 text-center">
                            <img src={logo3}/>
                        </Col>
                    </Row>
                    <Form name="normal_login" className="login-form">
                        <Form.Item name="username" rules={[{ required: true, message: 'Campo obbligatorio' }]}>
                            <Input name='userinput' size="large" prefix={<UserOutlined translate={undefined} className="site-form-item-icon" />}  placeholder="Username" allowClear onChange={this.onChangeUser}  />
                        </Form.Item>
                        <Form.Item name="password" rules={[{ required: true, message: 'Campo obbligatorio' }]}>
                            <Input size="large" prefix={<LockOutlined translate={undefined} className="site-form-item-icon" />} type="password" placeholder="Password" onChange={this.onChangePassword}/>
                        </Form.Item> 
                        <Form.Item>
                            <Button type="primary" className="login-form-button" onClick={() => this.onLogin(userlogin, passlogin)}>
                                Login
                            </Button>
                        </Form.Item>
                    </Form>
                </Modal>
            </>
        )
    }

}