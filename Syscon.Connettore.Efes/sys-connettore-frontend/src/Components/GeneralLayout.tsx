import React, { Component } from 'react'

import 'antd/dist/antd.css';
import '../App.css';

import { Alert, Avatar, Button, Col, Form, Input, Layout, Menu, Modal, Row, Tabs } from 'antd';
import EuroOutlined from '@ant-design/icons/EuroOutlined';
import { HomeOutlined, IdcardOutlined, LockOutlined, LogoutOutlined, QuestionCircleOutlined, SettingOutlined, TagsOutlined, ToolOutlined, UserOutlined } from '@ant-design/icons';
import SubMenu from 'antd/lib/menu/SubMenu';

import logo2 from '../Images/logo-sysconanalytics-b.png'
import logo3 from '../Images/logo-sysconanalytics.png'

import { Link } from 'react-router-dom';
import Home from './Home';
import axios from 'axios';

const { Header, Content, Footer, Sider } = Layout
const { TabPane } = Tabs;

interface LoginUser {
    id: string
}


export default class GeneralLayout extends Component<any, any> {

    constructor(props: any) {
        super(props);


        const panes = new Array(1).fill(null).map((_, index) => {
            const id = String(index + 1)
            return { title: `Home`, content: <Home></Home>, key: id };
          });

        this.state = { 
            login: false,
            loading: false,
            user: [],
            userlogin: '',
            passlogin: '',
            loginVisible: true,
            appVisible: 'hidden',
            newTabIndex : 0,
            collapsed: true,
            activeKey: panes[0].key,
            panes,
            logoutKeys: '',
            showAlert:  false,
            setShowAlert: false
        }


      }
    
    // state = {
    //     collapsed: true,
    //   };

    onCollapse = collapsed => {
        console.log(collapsed);
        this.setState({ collapsed });
      };

    onChange = activeKey => {
    this.setState({ activeKey });
    };

    onEdit = (targetKey, action) => {
    this[action](targetKey);
    };

    add = (title, newTabIndex, content) => {
        const { panes } = this.state;
        const activeKey = newTabIndex;
        panes.push({ title: title, content: content, key: activeKey });
        this.setState({ panes, activeKey, newTabIndex });
      };

    remove = targetKey => {
    let { activeKey } = this.state;
    let lastIndex;
    this.state.panes.forEach((pane, i) => {
        if (pane.key === targetKey) {
        lastIndex = i - 1;
        }
    });
    const panes = this.state.panes.filter(pane => pane.key !== targetKey);
    if (panes.length && activeKey === targetKey) {
        if (lastIndex >= 0) {
        activeKey = panes[lastIndex].key;
        } else {
        activeKey = panes[0].key;
        }
    }
    this.setState({ panes, activeKey });
    };

    onChangeUser = (e) => {
        this.setState({ userlogin: e.target.value })
    }

    onChangePassword = (e) => {
        this.setState({ passlogin: e.target.value })
    }


    onLogin = (username: string, password: string) => {
    
        this.setState({ showAlert : false, loading: true })
        const rlogin = axios.get('https://localhost:44369/api/v1/Login?username=' + username + '&password=' + password, { headers: {'Accept': 'application/json','Content-Type': 'application/json'}})
                            .then((response) => {
                                const login = response.data;
                                const user = response.data.item2;

                                this.setState({ login, user, loading: false })

                                if (this.state.login.item1 == true) {

                                    this.setState({ loginVisible : false, appVisible: 'visible'  })
                                    
                                }
                                else if (this.state.login.item1 == false) {
                                    
                                    this.setState({ showAlert : true })
                        
                                }
                            })
                            
        

    }

    onLogout = () => {
        this.setState({ loginVisible : true, appVisible: 'hidden', logoutKeys: ''  })
    }

    render() {

        let { loading, loginVisible, appVisible, collapsed, newTabIndex, userlogin, passlogin, user, logoutKeys, showAlert } = this.state;

        return (
            <>
            <Layout  >
                <Header style={{position: 'fixed', zIndex: 1, width: '100%', visibility: appVisible}} >
                    <div style={{width: '100%'}}  >
                        <img src={logo2} 
                        //width="300" 
                        //height="100"
                            />
                        <Menu theme='dark' mode="horizontal" style={{ marginRight: '20px', float: 'right'}} selectedKeys={logoutKeys}  >
                            <Menu.Item key="1">
                                <Avatar icon={<UserOutlined style={{ verticalAlign: 4 }} translate={undefined}/>}/>
                                &nbsp;
                                &nbsp;
                                {user.syS_DESCRIZIONE}
                            </Menu.Item>
                            <Menu.Item key="2"><QuestionCircleOutlined translate={undefined} /></Menu.Item>
                            <Menu.Item key="3" onClick={() => this.onLogout()}><LogoutOutlined translate={undefined} /></Menu.Item>
                        </Menu>
                    {/* </div> */}

                    </div>
                </Header>
                <Layout style={{marginTop: 70, visibility: appVisible }}>
                    <Sider width={250} className="site-layout-background" collapsed={true} style={{overflow: 'auto', height: '100vh', position: 'fixed', left: 0}}>   
                        <Menu theme='light' mode="inline" defaultSelectedKeys={['H1']} style={{ height: '100%', borderRight: 0 }}>
                            <Menu.Item key="H1" icon={<HomeOutlined translate={undefined}/>}>
                                <Link to="/" >Home</Link>
                            </Menu.Item>
                            <SubMenu key="sub2" title="Risorse" disabled={true} icon={<IdcardOutlined translate={undefined}/>}>
                                <Menu.Item key="R1">
                                    <Link to="/risorsedashboard">Dashboard risorse</Link>
                                </Menu.Item>
                                <Menu.Item key="R2">
                                    <Link to="/risorse">Elenco risorse</Link>
                                </Menu.Item>
                            </SubMenu>
                            <Menu.Item key="U1" disabled={true} icon={<UserOutlined translate={undefined}/>} >
                                <Link to="/" >Utenti</Link>
                            </Menu.Item>
                        </Menu>
                    </Sider>
                    <Layout style={{ padding: '0 24px 24px 94px' }}>
                        <Tabs hideAdd onChange={this.onChange} activeKey={this.state.activeKey} type="editable-card" onEdit={this.onEdit} style={{ padding: 24, margin: 0, minHeight: 840 }}>
                            {this.state.panes.map(pane => (
                                <TabPane tab={pane.title} key={pane.key}>
                                {pane.content}
                                </TabPane>
                            ))}
                        </Tabs>
                        <Footer style={{ textAlign: 'center', marginBottom: 0 }} >Syscon Connettore ©2022</Footer>
                    </Layout>
                </Layout>
            </Layout>

            {/* Modale per la login*/}
            <Modal visible={loginVisible} closable={false} footer={null}>
                <Row>
                    <Col className="col-12 text-center" >
                        <img src={logo3} width={350} height={165} />
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
                        <Button type="primary" style={{ float: 'right'}} loading={loading} onClick={() => this.onLogin(userlogin, passlogin)}>
                            Login
                        </Button>
                    </Form.Item>
                </Form>
                {showAlert&&
                <Alert message="Login fallito" type="error" showIcon closable />
                }

            </Modal>
            </>
        )
    }




}