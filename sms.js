const way2sms = require('way2sms');
// way2sms.reLogin(<mobileno>, <password>): returns login cookie (promise)
// way2sms.smstoss(<cookie>, <tomobile>, <message>): sends sms (promise)
 
cookie = way2sms.login('9561375642', 'madhvi12345'); // reLogin
// <cookie string>
setTimeout(function(){
  way2sms.send(cookie, '9967569282', 'good morning'); // smstoss
},100000*60);
 