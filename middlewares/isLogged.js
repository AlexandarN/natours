     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
const jwt = require('jsonwebtoken');


	// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
const catchAsync = require('../utils/catchAsync');
const AppError = require('../utils/AppError.js');
const env = require('../config/env');
const User = require('../models/User');

	// 3) CONSTANTs	

     // 4) MIDDLEWAREs	
module.exports = catchAsync(async (req, res, next) => {
     // 1. Get the Token
     let token;
     if(req.headers.authorization) {
          token = req.headers.authorization.split(' ')[1];
     }
          // Check if token exists?
     if(!token) {
          return next(new AppError('Please log in to access this page!', 401));
     }
     // 2. Validate the token (verification)
     const decodedToken = jwt.verify(token, env.JWT_SECRET);  // decodedToken = verified token
     if(!decodedToken) {
		return next(new AppError('Invalid token. Please log in again!', 401));
     }
     // 3. Check if the user wanting to log in still exists in the DB
     const user = await User.findById(decodedToken.id);
	if(!user) {							               
		return next(new AppError('The user attached to this token does no longer exist.', 401));
     }
     // 4. Check if the user has changed his password since the token was issued
     if(user.passwordChangedAt) {
		const passChangeTime = parseInt(user.passwordChangedAt.getTime() / 1000, 10);  
		const tokenIssueTime = decodedToken.iat;					 
		if(tokenIssueTime < passChangeTime) {               
			return next(new AppError('User recently changed password. Please log in again!', 401));
		}
     }
	// 5. Grant access to the protected route + add user to the arrived request
	req.user = user;
     next();
})