from win32com.client import Dispatch
from flask import Flask, request, jsonify
import sys
import pythoncom
from custom_logger import logger


app = Flask(__name__)

@app.route('/send/outlook/mail/', methods=['POST'])
def send_outlook_mail():
    
    try:
        print('**************')
        outlook_client = Dispatch('Outlook.Application', pythoncom.CoInitialize())
        print('++++++++++++++')
        
        pay_load = request.json
        logger.debug(pay_load)
        
        mail = outlook_client.CreateItem(0)
        mail.To = ';'.join(pay_load['to']) if type(pay_load['to']) == list else pay_load['to']
        mail.Subject = pay_load['subject']
        mail.CC = ';'.join(pay_load['cc']) if type(pay_load['cc']) == list else pay_load['cc']
        mail.BCC = ';'.join(pay_load['bcc']) if type(pay_load['bcc']) == list else pay_load['bcc']
        print(pay_load)
        
        if pay_load.get('logo'):
            try:
                logo_attach = mail.Attachments.Add(pay_load['logo'])
                logo_attach.PropertyAccessor.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', 'logo')
            except Exception as ex:
                logger.error((type(ex), sys.exc_info()[1], sys.exc_info()[2]))
            
        with open(pay_load['body'], mode='r') as f:
            _body = f.read()
        if pay_load['body_type'] == 'html':
            mail.HTMLBody = _body
        else:
            mail.Body = _body
            
        if pay_load.get('attachments'):
            try:
                if type(pay_load['attachments']) == list:
                    for attach in pay_load['attachments']:
                        try:
                            mail.Attachments.Add(attach)
                        except Exception as ex:
                            logger.error((type(ex), sys.exc_info()[1], sys.exc_info()[2]))
                else:
                    mail.Attachment.Add(pay_load['attachments'])
            except Exception as ex:
                logger.error((type(ex), sys.exc_info()[1], sys.exc_info()[2]))
                
        mail.Send()
    except Exception as ex:
        logger.error((type(ex), sys.exc_info()[1], sys.exc_info()[2]))
        return jsonify({'message': str((type(ex), sys.exc_info()[1], sys.exc_info()[2]))})
        
    else:
        logger.info('mail sent successfully')
        return jsonify({'message': 'mail sent successfully'})
        
# @app.route('/', methods=['GET'])
# def test_endpoint():
    # return jsonify({'message': 'You are in OutLookMail_Server_UsingWIN32_Flask-Directory'})
     
     
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
