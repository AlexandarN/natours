// 1)  IMPORTING NPM PACKAGEs and NODE MODULEs	
const nodemailer = require('nodemailer');
// const sendgrid = require('nodemailer-sendgrid-transport');


// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
const env = require('../config/env');

// 3) CONSTANTs	


// 4) MIDDLEWAREs	
module.exports = async options => {
     // 1. Create transporter
     const transporter = nodemailer.createTransport({
          host: env.EMAIL_HOST,						         
		port: env.EMAIL_PORT,
		auth: {
			user: env.EMAIL_USERNAME,
               pass: env.EMAIL_PASSWORD 
          }
     });
     // 2. Send email
	await transporter.sendMail({                      
		to: options.email,
		from: 'Aleksandar Nikolic <alex@natours.com>',
		subject: options.subject,
		text: options.message
	}); 
}

