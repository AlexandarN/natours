	// 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
const express = require('express');  
const morgan = require('morgan');
const path = require('path');

	// 1.1)  DEPLOYMENT PACKAGEs
const rateLimit = require('express-rate-limit');
const helmet = require('helmet');
const mongoSanitize = require('express-mongo-sanitize');
const xss = require('xss-clean');
const hpp = require('hpp');

	// 2) IMPORTING our custom files (ROUTEs, CONTROLLERs, MODELs, ...)	
const toursRoutes = require('./routes/toursRoutes');
const usersRoutes = require('./routes/usersRoutes');
const reviewsRoutes = require('./routes/reviewsRoutes');

const AppError = require('./utils/AppError');   
const errorController = require('./controllers/errorController');


	// 3) CONSTANTs
	// 3.1)  Express initatation			
const app = express();	

	// 3.2)  Consecutive Http requests limiter		
const limiter = rateLimit({
	max: 100,    // max 100 requests from a single IP address in an hour
	windowsMs: 60 * 60 * 1000,    // in miliseconds
	message: 'Too many requests from this IP, please try again in an hour!'
}); 

	// 4) VIEW ENGINE setting	

	// 5) MIDDLEWAREs
	// 5.1) MIDDLEWAREs for setting a static PUBLIC folder  
app.use(express.static(path.resolve('public')));        

  // 5.2) PARSING MIDDLEWAREs for POST request inputs
app.use(express.json());

	// 5.3) SESSION and FLASH MIDDLEWAREs	

	// 5.4) USER AUTHENTICATION  MIDDLEWAREs
		// 1st way - 'PASSPORT' Authentication Middlewares (this must go above res.locals.user)
		// 2nd way - Midd. for catching logged in user from session and putting it into every http request - we can then use it in Controllers


	// 5.5)  GLOBAL VARIABLEs MIDDLEWAREs
		// 5.5.1) GLOBAL variables MIDDLEWARE (app.locals) - for catching DB RESOURCES in ALL RESPONSES
		// 5.5.2) GLOBAL variables MIDDLEWARE (res.locals) - using SESSION and FLASH - we can then use these variables in all views (responses)	


	// 5.6)  CSRF SETUP and CSRF MIDDLEWAREs                                       
		// 5.6.1) ROUTES that we want to be ignored by CSRF Middleware we need to set them above the CSRF Midd. function
		// 5.6.2) CSRF MIDDLEWARE - ROUTES below will be affected by CSRF Midd.
		// 5.6.3) GLOBAL variables MIDDLEWARE (res.locals) for CSRF TOKEN


	// 5.7) MIDDLEWAREs for LOGGING
if(process.env.NODE_ENV !== 'production') {
     app.use(morgan('dev'));
}


	// 5.8)  SECURITY MIDDLEWAREs (to be used for DEPLOYMENT)
		// 5.8.1)  Middl for LIMITATION of a MAX NUMBER of CONSECUTIVE HTTP REQUESTs (BRUTE FORCE ATTACKS)
app.use('/', limiter);      
		// 5.8.2)  SETTING SECURITY HTTP HEADERS
app.use(helmet());
		// 5.8.3) PARSED DATA SANITIZATION against NOSQL query ATTACKS
app.use(mongoSanitize());	                                   
		// 5.8.4) PARSED DATA SANITIZATION against XSS (CROSS_SITE_SCRIPTING) ATTACKS
app.use(xss());	
		// 5.8.5) REQ.QUERY PARAMETER POLUTION PREVENTION (from entering duplicate parameters in req.query - which have strings as values, e.g. .../tours?sort=duration&sort=price (duration and price are strings, so it will return an error)
app.use(hpp({
	whitelist: ['duration', 'price', 'difficulty', 'maxGroupSize', 'ratingsAverage', 'ratingsQuantity']
})); 
		// 5.8.6)  COMPRESSION of ASSET FILES (CSS and JS, not including IMAGEs)

		// 5.8.7)  Reading of HTTPS files for SSL/TLS protection - this is very optional (normally hosting providers do this)		



     // 5.9) ROUTES MIDDLEWAREs         
app.use('/tours', toursRoutes);	
app.use('/users', usersRoutes);	
app.use('/reviews', reviewsRoutes);	
app.use('*', (req, res, next) => {
	next(new AppError(`Can't find '${req.originalUrl}' route on this server!`, 404));     
});	


	// 5.10) ERROR handling MIDDLEWARE           
app.use(errorController);	
	

	// 6) EXPORTING APP			 
module.exports = app;						                                                               	
