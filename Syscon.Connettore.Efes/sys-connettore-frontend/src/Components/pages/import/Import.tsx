import { Badge, Button, Card, Col, ConfigProvider, DatePicker, Descriptions, Divider, Modal, Progress, Radio, Row, Space, Spin, Statistic, Steps, Table, Tabs, Tooltip } from 'antd';
import axios from 'axios';
import React, { Component, useState, useEffect } from 'react'
import 'moment/locale/it';
import { __values } from 'tslib';

const gridStyle = {
    width: '25%',
    textAlign: 'center',
  };

interface ImportState {

    pn : any,
    radiomode :any,

    loading: boolean,
    filteredInfo: any,
    sortedInfo: any,
    current: number,
    setCurrent: number
}

export default class Import extends Component<any, ImportState> {


    constructor(props: any) {
        super(props);

        this.state = {  pn: [],
                        radiomode: '',

                        loading: true,
                        filteredInfo: null, 
                        sortedInfo: null,  
                        current: 0,
                        setCurrent: 0

                     } 
    }

    componentDidMount() {

        const rpn = axios.get('https://localhost:44327/api/v1/ImportPN', { headers: {'Accept': 'application/json','Content-Type': 'application/json'}})

        axios.all([ rpn  ])
        .then(axios.spread((...responses)  => {
            const pn = responses[0].data;

            this.setState({     pn,
                                loading : false });
        }
        ))
    }


    handleChange = (pagination, filters, sorter) => {
        console.log('Various parameters', pagination, filters, sorter);
        this.setState({
            filteredInfo: filters,
            sortedInfo: sorter,
        });
    };


    render() {
  
    let {   pn,
            radiomode, 
            sortedInfo, 
            filteredInfo, 
            loading,
            current,
            setCurrent} = this.state;

    sortedInfo = sortedInfo || {};
    filteredInfo = filteredInfo || {};


    const { TabPane } = Tabs;
    const { Step } = Steps;

    const steps = [
        {
          title: 'Elabora',
          content: 'Lettura dati',
        },
        {
          title: 'Verifica',
          content: 'Verifica dati formali',
        },
        {
          title: 'Importa',
          content: 'Importazione dati',
        },
      ];

    const colPN = [
            {
                title: 'Esito',
                dataIndex: '',
                key: 'esito',
                render: ({esito}) => {
                    let progresscircle;
                    if (esito === 0) {
                        progresscircle = <Progress type="circle" percent={0} width={40} format={() => ''} />
                    }
                    return progresscircle;
                },
                sorter: (a, b) => a.esito - b.esito,
                sortOrder: sortedInfo.columnKey === 'esito' && sortedInfo.order,
                width: 50
            },
            {
                title: 'Ditta',
                dataIndex: 'ditta',
                key: 'ditta',
                filteredValue: filteredInfo.ditta || null,
                onFilter: (value, record) => record.ditta.includes(value),
            },
            {
                title: 'Flusso',
                dataIndex: '',
                key: 'goKey',
                render: ({goKey}) => {
                    return <b>{goKey}</b>
                },
                filteredValue: filteredInfo.goKey || null,
                onFilter: (value, record) => record.goKey.includes(value),
            },
   ]

       return (
           <>
            <Card>
                    {/* <div style={{marginBottom: 65}}>
                    <Title level={2}>Webrecall</Title>
                    </div> */}
                    <div className="site-card-wrapper">
                        <Row gutter={16}>
                            <Col span={3}></Col>
                            <Col span={18}>
                                <Steps current={current} type="navigation">
                                    {steps.map(item => (<Step key={item.title} title={item.title} description={item.content} />))}
                                </Steps>
                            </Col>
                            <Col span={3}></Col>
                        </Row>
                        <br></br>
                        <br></br>
                        
                        <Tabs defaultActiveKey="1" tabPosition='top' type="card">
                            <TabPane tab="Clienti" key="1">
                                <Row gutter={16}>
                                    {/* <Col span={3}>
                                    </Col> */}
                                    <Col span={24}>
                                            <Spin spinning={loading}>
                                                <Table columns={colPN} dataSource={pn} size="small" bordered={true} pagination={false} onChange={this.handleChange}/>
                                            </Spin>
                                    </Col>
                                </Row>
                            </TabPane>
                            <TabPane tab="Documenti" key="2">
                                <Row gutter={16}>
                                    <Col span={21}>
                                        <Spin spinning={loading}>
                                            {/* <Table columns={colOperatore} dataSource={wrop} size="middle" bordered={true} pagination={false} onChange={this.handleChange}                                         
                                                    summary={pageData => {
                                                    let totaltsent = 0;
                                                    let totaltsa = 0;
                                                    let totaltss = 0;
                                                    let totaltssap = 0;
                                                    let totaltssred = 0;
                                                    let totaltshr = 0;
                                                    let totaltscrm = 0;
                                                    let totalmyt = 0;
                                                    let totalop = 0;

                                                    pageData.forEach(({ wR_TSENT, 
                                                                        wR_TSA,
                                                                        wR_TSS,
                                                                        wR_TSSAP,
                                                                        wR_TSSRED,
                                                                        wR_TSHR,
                                                                        wR_TSCRM,
                                                                        wR_MYT,
                                                                        wR_TOTALE}) => {
                                                    totaltsent += wR_TSENT;
                                                    totaltsa += wR_TSA;
                                                    totaltss += wR_TSS;
                                                    totaltssap += wR_TSSAP;
                                                    totaltssred += wR_TSSRED;
                                                    totaltshr += wR_TSHR;
                                                    totaltscrm += wR_TSCRM;
                                                    totalmyt += wR_MYT;
                                                    totalop += wR_TOTALE;
                                                    });

                                                    return (
                                                    <>
                                                        <Table.Summary.Row >
                                                            <Table.Summary.Cell index={0} >
                                                                <div style={{textAlign: 'left'}} >
                                                                    <b>Totali</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={1}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltsent}</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={2}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltsa}</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={3}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltss}</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={4}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltssap}</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={5}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltssred}</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={6}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltshr}</b>
                                                                </div>                                                           
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={7}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltscrm}</b>
                                                                </div>                                                           
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={8}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totalmyt}</b>
                                                                </div>                                                           
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={9}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totalop}</b>
                                                                </div>                                                           
                                                            </Table.Summary.Cell>
                                                        </Table.Summary.Row>
                                                    </>
                                                    );
                                                }}/> */}
                                        </Spin>
                                    </Col>
                                </Row>
                            </TabPane>
                            <TabPane tab="PrimaNota" key="3">
                                <Row gutter={16}>
                                    <Col span={21}>
                                        <Spin spinning={loading}>
                                            {/* <Table columns={colProdotto} dataSource={wrap} size="middle" bordered={true} pagination={{pageSize : 13}} onChange={this.handleChange}                                         
                                                    summary={pageData => {
                                                    let totaltnew = 0;
                                                    let totaltclose = 0;
                                                    let totaltopen = 0;
                                                    let totalap = 0;

                                                    pageData.forEach(({ wR_TNEW, 
                                                                        wR_TCLOSE,
                                                                        wR_TOPEN,
                                                                        wR_TOTALE}) => {
                                                    totaltnew += wR_TNEW;
                                                    totaltclose += wR_TCLOSE;
                                                    totaltopen += wR_TOPEN;
                                                    totalap += wR_TOTALE;
                                                    });

                                                    return (
                                                    <>
                                                        <Table.Summary.Row >
                                                            <Table.Summary.Cell index={0} >
                                                                <div style={{textAlign: 'left'}} >
                                                                    <b>TOTALI</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={1}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltnew}</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={2}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltclose}</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={3}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totaltopen}</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                            <Table.Summary.Cell index={4}>
                                                                <div style={{textAlign: 'right'}} >
                                                                    <b>{totalap}</b>
                                                                </div>
                                                            </Table.Summary.Cell>
                                                        </Table.Summary.Row>
                                                    </>
                                                    );
                                                }}/> */}
                                        </Spin>
                                    </Col>
                                    <Col span={3}>
                                    {/*     <Row>
                                            <Radio.Group  defaultValue='oggi' onChange={this.onChangeAP}>
                                                <Card title='Data riferimento' headStyle={{height: 10}} > 
                                                    <Space direction="vertical">
                                                        <Radio value="oggi">Oggi</Radio>
                                                        <Radio value="ieri">Ieri</Radio>
                                                        <Radio value="last7">Ultimi 7 gg</Radio>
                                                        <Radio value="last15">Ultimi 15 gg</Radio>
                                                        <Radio value="last30">Ultimi 30 gg</Radio>
                                                        <Radio value="totale">Totale</Radio>
                                                        <Radio value="datapers">Data personalizzata</Radio>
                                                        <DatePicker  placeholder="Da data" locale={locale} disabled={enableddate} onChange={(data) => this.onChangeAPDateStart(moment(data).format("YYYY/MM/DD"))} ></DatePicker>
                                                        <DatePicker placeholder="A data" locale={locale} disabled={enableddate} onChange={(data) => this.onChangeAPDateFinish(moment(data).format("YYYY/MM/DD"))} ></DatePicker>
                                                        <Button type="primary" disabled={enableddate} onClick={() => this.onChangeAPDate(datestart, datefinish)} >Applica</Button>
                                                    </Space>
                                                </Card>
                                            </Radio.Group>
                                        </Row>
                                        {/* <br></br>
                                        <Row>
                                            <Radio.Group  defaultValue='totali' onChange={this.onChange}>
                                                <Card title='Modalità' headStyle={{height: 10}} > 
                                                    <Space direction="vertical">
                                                        <Radio value="totali" >Totali</Radio>
                                                        <Radio value="media" >Media</Radio>
                                                    </Space>
                                                </Card>
                                            </Radio.Group>
                                        </Row> */}
                                    </Col>
                                </Row>

                            </TabPane>
                        </Tabs>     
                    </div>
            </Card>

            <Modal></Modal>                            


           </>
       );
    }
    



}