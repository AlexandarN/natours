     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
	
	// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
	const env = require('../config/env');      			

	// 3) CONSTANTs	

	
	// 4) CONTROLLERs	
module.exports = (err, req, res, next) => {   
     // Define statusCode and status	
	err.statusCode = err.statusCode || 500; 
	err.status = err.status || 'error';				         
	// SEND RESPONSE to FE when in DEVELOPMENT mode
     if(env.NODE_ENV === 'development') {
		// 1.1 If there is an error based on a wrong req.params parameter sent in Url
		if(err.name === 'CastError') {		
			const message = `Invalid ${err.path}: ${err.value}`;		
			res.status(400).json({        		
				status: 'fail',
				error: err,
				message: message,          		   
		          stack: err.stack 
			});
		// 1.2 If there is an error based on a duplicate value entered for a unique property of a Schema modeled object
		} else if(err.name === 'MongoError') {		
			const value = err.errmsg.match(/(["'])(\\?.)*?\1/)[0];
			const message = `Duplicate field value entered: ${value}`;		
			res.status(400).json({        		
				status: 'fail',
				error: err,
				message: message,          		   
		          stack: err.stack 
			});
		// 1.3 If there are validation errors based on the incorrect values entered for a certain properties of a modeled object
		} else if(err.name === 'ValidationError') {		
			const errorMsgs = Object.values(err.errors).map(prop => prop.message);
			const message = `Invalid input data: ${errorMsgs.join('. ')}`;	
			res.status(400).json({        		
				status: 'fail',
				error: err,
				message: message,          		   
				stack: err.stack 
			});	
		// 1.4 If there is an invalid token error - when a user wants to login with incorrect token sent in the Authorization Header of http request
		} else if(err.name === 'JsonWebTokenError') {		
			const message = 'Invalid token. Please log in again!';	
			res.status(401).json({        		
				status: 'fail',
				error: err,
				message: message,          		   
				stack: err.stack 
			});	
		// 1.5 If there is an expired token error - when a user wants to login with an expired token sent in the Authorization Header of http request
		} else if(err.name === 'TokenExpiredError') {		
			const message = 'Your token has expired. Please log in again!';	
			res.status(401).json({        		
				status: 'fail',
				error: err,
				message: message,          		   
				stack: err.stack 
			});	
		// For all other known or unknown errors	
		} else {
			res.status(err.statusCode).json({        		
				status: err.status,
				error: err,
				message: err.message,          		   
				stack: err.stack 
			});
		}	
	// SEND RESPONSE to FE when in PRODUCTION mode
	} else if(env.NODE_ENV === 'production') {
		// 1. SEND message to the client - when the error is KNOWN 
			// 1.1 If there is an error based on a wrong req.params parameter sent in Url
		if(err.name === 'CastError') {		
			const message = `Invalid ${err.path}: ${err.value}`;		
			res.status(400).json({        		
				status: 'fail',
				message: message    		   
			});
			// 1.2 If there is an error based on a duplicate value entered for a unique property of a modeled object
		} else if(err.name === 'MongoError') {		
			const value = err.errmsg.match(/"([^"]*)"/)[1];
			const message = `Duplicate field value entered: ${value}`;		
			res.status(400).json({        		
				status: 'fail',
				message: message   		   
			});
			// 1.3 If there are validation errors based on the incorrect values entered for a certain properties of a modeled object
		} else if(err.name === 'ValidationError') {		
			const errorMsgs = Object.values(err.errors).map(prop => prop.message);
			const message = `Invalid input data: ${errorMsgs.join('. ')}`;	
			res.status(400).json({        		
				status: 'fail',
				message: message        		   
			});	      
			// 1.4 If there is an error anticipated in controllers - e.g. could not find the requested document in DB
		} else if(err.isOperational) {				           
			res.status(err.statusCode).json({        		
				status: err.status,
                    message: err.message 
			});       
			// 1.5 If there is an invalid token error - when a user wants to login with incorrect token sent in the Authorization Header of http request
		} else if(err.name === 'JsonWebTokenError') {		
			const message = 'Invalid token. Please log in again!';	
			res.status(401).json({        		
				status: 'fail',
				message: message          		   
			});	
			// 1.6 If there is an expired token error - when a user wants to login with an expired token sent in the Authorization Header of http request
		} else if(err.name === 'TokenExpiredError') {		
			const message = 'Your token has expired. Please log in again!';	
			res.status(401).json({        		
				status: 'fail',
				message: message         		   
			});	
		} else {
		// 2. DON'T leak the message to the client - when the error is UNKNOWN (programming or other error)
			// 2.1) Log the error
			console.error('ERROR💥: ', err);
			// 2.2) Send response
			res.status(500).json({        		
				status: 'error',
                    message: 'Something went wrong!' 
               });              
		}
	}
}

               