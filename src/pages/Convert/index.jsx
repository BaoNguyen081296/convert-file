import React from 'react';
import './Convert.scss';
import { Form, Button, Upload, notification } from 'antd';
import { UploadOutlined } from '@ant-design/icons';
import { exportFile } from 'export/json';
const typeXLSX = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

const title = 'Please import the XLSX file.';
function index() {
  const normFile = e => {
    if (e && typeXLSX.indexOf(e.file.type) === -1 && e.file.response) {
      notification['error']({
        duration: 5,
        message: 'Invalid file input',
        description: title
      });
      return;
    }
    if (Array.isArray(e)) return e;
    return e && e.fileList;
  };
  const handleSubmit = e => {
    exportFile(e.file[0]);
  };
  const propsUpload = {
    name: 'file',
    listType: 'text',
    maxCount: 1,
    accept: typeXLSX
  };

  return (
    <div className="_convert">
      <div className="_convert-content">
        <h1>Convert to JSON</h1>
        <h5>{title}</h5>
        <Form className="_convert-content-form" onFinish={handleSubmit} initialValues={{}}>
          <Form.Item
            name="file"
            label="Upload"
            valuePropName="fileList"
            getValueFromEvent={normFile}
            rules={[
              {
                required: true,
                message: title
              }
            ]}
          >
            <Upload {...propsUpload}>
              <Button icon={<UploadOutlined />}>Click to upload</Button>
            </Upload>
          </Form.Item>
          <Form.Item
            wrapperCol={{
              span: 12,
              offset: 6
            }}
          >
            <Button type="primary" htmlType="submit">
              Convert
            </Button>
          </Form.Item>
        </Form>
      </div>
    </div>
  );
}
export default index;
