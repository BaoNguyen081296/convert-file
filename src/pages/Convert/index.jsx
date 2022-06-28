import React, { useState, useMemo, useRef } from 'react';
import './Convert.scss';
import { Form, Button, Upload, notification, Select } from 'antd';
import { UploadOutlined } from '@ant-design/icons';
import { exportFile } from 'export/json';
import { DESTINATION_FILE, TYPE } from 'utils/utils';

const { Option } = Select;

function Convert() {
  const [type, setType] = useState(TYPE.TO_XLSX);
  const [destination, setDestination] = useState(DESTINATION_FILE.TO_FIRST);
  const formRef = useRef(null);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const typeXLSX =
    type === TYPE.TO_JSON
      ? ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']
      : ['application/json'];
  const title = `Please import the ${type === TYPE.TO_JSON ? 'XLSX' : 'JSON'} file.`;

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
  const handleSubmit = ({ file }) => {
    exportFile({ type, file, destination });
  };
  const onTypeChange = e => {
    setType(e);
    formRef.current.validateFields();
  };

  const onDesChange = e => {
    setDestination(e);
  };

  const propsUpload = useMemo(
    () => ({
      name: 'file',
      listType: 'text',
      maxCount: 2,
      accept: typeXLSX
    }),
    [typeXLSX]
  );

  return (
    <div className="_convert">
      <div className="_convert-content">
        <h1>{type === TYPE.TO_JSON ? 'Convert XLSX to JSON' : 'Convert JSON to XLSX'}</h1>
        <h5>{title}</h5>
        <Form
          ref={formRef}
          className="_convert-content-form"
          onFinish={handleSubmit}
          initialValues={{ type, destination }}
        >
          <Form.Item name="type" label="Type">
            <Select onChange={onTypeChange}>
              <Option value={TYPE.TO_JSON}>XLSX to JSON</Option>
              <Option value={TYPE.TO_XLSX}>JSON to XLSX</Option>
            </Select>
          </Form.Item>
          {type === TYPE.TO_JSON && (
            <Form.Item name="destination" label="Destination">
              <Select onChange={onDesChange}>
                <Option value={DESTINATION_FILE.TO_FIRST}>To First Language</Option>
                <Option value={DESTINATION_FILE.TO_SECOND}>To Second language</Option>
              </Select>
            </Form.Item>
          )}
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
export default Convert;
