import React, { Component } from 'react'

import 'antd/dist/antd.css';
import '../App.css';

import { Alert, Avatar, Button, Card, Col, Descriptions, Divider, Dropdown, Form, Input, Layout, Menu, Modal, Row, Space, Tabs } from 'antd';
import EuroOutlined from '@ant-design/icons/EuroOutlined';
import { DownloadOutlined, HomeOutlined, IdcardOutlined, LockOutlined, LogoutOutlined, QuestionCircleOutlined, SettingOutlined, ShopOutlined, TagsOutlined, ToolOutlined, UserOutlined } from '@ant-design/icons';
import SubMenu from 'antd/lib/menu/SubMenu';

import logo2 from '../images/logo-sysconanalytics-b.png'
import logo3 from '../images/logo-sysconanalytics.png'

import { Link } from 'react-router-dom';
import Home from './Home';
import axios from 'axios';
import Import from './pages/import/Import';

const { Header, Content, Footer, Sider } = Layout
const { TabPane } = Tabs;

interface LoginUser {
    id: string
}


export default class Container extends Component<any, any> {

    constructor(props: any) {
        super(props);


        const panes = new Array(1).fill(null).map((_, index) => {
            const id = String(index + 1)
            return { title: <div><HomeOutlined style={{ verticalAlign: 2 }} translate={undefined}/>Dashboard</div>, content: <Home></Home>, key: id };
          });

        this.state = { 
            login: true,
            loading: false,
            user: [],
            userlogin: '',
            passlogin: '',
            loginVisible: false,
            appVisible: 'visible',
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

        const userdescr = (
            <Card>
                <Descriptions bordered column={1} >
                    <Descriptions.Item label="Username" children={['admin']} ></Descriptions.Item>
                    <Descriptions.Item label="Nome" children={['Administrator']} ></Descriptions.Item>
                    <Descriptions.Item label="Email" children={['admin@syscon.it']} ></Descriptions.Item>
                </Descriptions>
            </Card>
        );

        const company = (
            <Menu theme='light' mode="horizontal" defaultSelectedKeys={['1']}>
                <Menu.Item key="1">
                    <Avatar icon={<ShopOutlined style={{ verticalAlign: 4 }} translate={undefined}/>}/>
                    &nbsp;
                    &nbsp;
                    Federazione Italiana Tennis
                    {user.syS_DESCRIZIONE}
                </Menu.Item>
                <Menu.Item key="2">
                    <Avatar icon={<ShopOutlined style={{ verticalAlign: 4 }} translate={undefined}/>}/>
                    &nbsp;
                    &nbsp;
                    Syscon
                    {user.syS_DESCRIZIONE}
                </Menu.Item>
            </Menu>
        );

        
        return (
            <>
            <Layout  >
                <Header style={{position: 'fixed', zIndex: 1, width: '100%', visibility: appVisible, display: 'flex'}} >
                    <div style={{width: '50%'}}  >
                        <img src={logo2} 
                        //width="300" 
                        //height="100"
                            />
                    </div>
                    <div style={{width: '50%', justifyContent: 'flex-end', display: 'flex'}}>
                        <Dropdown overlay={userdescr}>
                            <div>
                                <Avatar icon={<UserOutlined style={{ verticalAlign: 4 }} translate={undefined}/>}/>
                                &nbsp;
                                &nbsp;
                                <a style={{ color: 'white' }} >Administrator</a>
                                {user.syS_DESCRIZIONE}
                                &nbsp;
                                &nbsp;
                            </div>
                        </Dropdown>
                        &nbsp;
                        &nbsp;
                        <Dropdown overlay={company}>
                            <div>
                                    <Avatar icon={<ShopOutlined style={{ verticalAlign: 4 }} translate={undefined}/>}/>
                                    &nbsp;
                                    &nbsp;
                                    <a style={{ color: 'white' }} >Federazione Italiana Tennis</a>
                                    
                            </div>
                        </Dropdown>
                        &nbsp;
                        &nbsp;
                        <Menu theme='dark' mode="horizontal" selectedKeys={logoutKeys}>
                            <Menu.Item key="3"><QuestionCircleOutlined translate={undefined} /></Menu.Item>
                            <Menu.Item key="4" onClick={() => this.onLogout()}><LogoutOutlined translate={undefined} /></Menu.Item>
                        </Menu>
                    </div>
                    {/* </div> */}


                </Header>
                <Layout style={{marginTop: 70, visibility: appVisible, marginBottom: 70 }}>
                    <Sider width={250} className="site-layout-background" collapsed={false} style={{overflow: 'auto', height: '100vh', position: 'fixed', left: 0}}>   
                        <Menu theme='light' mode="inline" defaultSelectedKeys={['H1']} style={{ height: '100%', borderRight: 0 }}>
                            <Menu.Item key="H1" icon={<HomeOutlined translate={undefined}/>}>
                                <Link to="/" >Dashboard</Link>
                            </Menu.Item>
                            <SubMenu key="sub2" title="Import" disabled={false} icon={<IdcardOutlined translate={undefined}/>}>
                                <Menu.Item key="R1">
                                    <Link onClick={() => this.add(<div><DownloadOutlined translate={undefined} style={{verticalAlign: 2}}/>Import movimenti</div>, `newTab${newTabIndex + 1}`, <Import></Import>)} to="/" >Import movimenti</Link>
                                </Menu.Item>
                            </SubMenu>
                            <Menu.Item key="U1" disabled={true} icon={<UserOutlined translate={undefined}/>} >
                                <Link to="/" >Utenti</Link>
                            </Menu.Item>
                        </Menu>
                    </Sider>
                    <Layout style={{ padding: '0 24px 24px 270px' }}>
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