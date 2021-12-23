     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	


	// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
const User = require('../models/User');
const catchAsync = require('../utils/catchAsync');
const AppError = require('../utils/AppError');
const handlerFactory = require('./handlerFactory');


	// 3) CONSTANTs	
	// 3.1) Function that enables to choose which parts of the req.body object (coming from Post request) to accept (which key + value pairs are allowed in req.body) based on the provided array of properties (keys), i.e. "...allowedProperties" -> this is needed in the 'updateMe' controller for security reasons (as we 	don't want a user to be able to change his 'role', for example)
const filterReqBody = (reqBody, ...allowedProperties) => {
	const newObj = {};
	Object.keys(reqBody).forEach(key => {	        
		if(allowedProperties.includes(key)) {
			newObj[key] = reqBody[key]; 
		}
	});
	return newObj;
}
	

	// 4)  MIDDLEWAREs	
	// 4.1)  Get Me Middleware - this is the middleware that we created in order to be able to use handlerFactory function 'getDoc' for 'Get Me' route
exports.getMe = (req, res, next) => {	                
	// Instead of req.params.id use req.user.id -> to catch current user from DB
	req.params.id = req.user.id;
	next();				
}


	// 4) CONTROLLERs	
exports.getUsers = handlerFactory.getAll(User);

exports.addUser = handlerFactory.addDoc(User);

exports.getUser = handlerFactory.getDoc(User);

exports.editUser = handlerFactory.editDoc(User);

exports.deleteUser = handlerFactory.deleteDoc(User);


exports.updateMe = catchAsync(async (req, res, next) => {       
	// 1. Create error if user sends password data (tries to update password) as this only allowed to be done in the update password route 
     if(req.body.password || req.body.passwordConfirm) {
		return next(new AppError('This route is not for password updates. Use other route!', 400));
	}
	// 2. Filter out unwanted properties of a user, that are not allowed to be updated in this controller (e.g. role, ...)
	const filteredRBody = filterReqBody(req.body, 'name', 'email');
	// 3. Find the user based on the req.user (acquired in isLogged.js middleware) and update it's data
	const user = await User.findByIdAndUpdate(req.user.id, filteredRBody, { 
		new: true,			       
		runValidators: true 
	});          						                
	// 4. Send the response	
	res.status(200).json({        	
		status: 'success',	                      
		data: {
			user: user 
		}
	}); 
}); 


exports.deleteMe = catchAsync(async (req, res, next) => {       
	// 1. Find the user based on the req.user (acquired in isLogged.js middleware) and update it's data
	await User.findByIdAndUpdate(req.user.id, { active: false });
	// 2. Send the response			        
	res.status(204).json({
		status: 'success',	                      
		data: null
	}); 
}); 



