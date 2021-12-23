     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');
     

	// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
const User = require('../models/User');
const catchAsync = require('../utils/catchAsync');
const AppError = require('../utils/AppError');
const env = require('../config/env');
const sendEmail = require('../utils/sendEmail');

     
     // 3) CONSTANTs	
const createToken = id => {							 
     return jwt.sign(
          {id: id},								         
          env.JWT_SECRET,              	  	         
          {expiresIn: env.JWT_EXPIRATION} 				         
     );	                           
}


const sendResponseWithToken = (user, statusCode, res) => {
     // 1. Create jwt token
     const token = createToken(user._id);
     // 2. Insert token into cookie and attach cookie to the response	
     res.cookie('jwt', token, {       
          expires: new Date(Date.now() + env.JWT_COOKIE_EXPIRATION * 24*60*60*1000),  // to get miliseconds
          secure: env.JWT_COOKIE_SECURE, 
          httpOnly: true 
     });
     user.password = undefined;    // this is temporarely -> just not to be sent in response (we do not save it to DB)
     // 3. Send response data to FE
     res.status(statusCode).json({
          status: 'success',
          token: token,
          data: {
               user: user
          }
     })
}


     // 4) CONTROLLERs	

exports.signup = catchAsync(async (req, res, next) => {
     // Create new user and save it in DB
     const newUser = await User.create({   	   		   
          name: req.body.name,
          email: req.body.email,
          password: req.body.password,
          passwordConfirm: req.body.passwordConfirm,
          passwordChangedAt: req.body.passwordChangedAt,
          role: req.body.role
     });
     // Create login token and send response to FE	
	sendResponseWithToken(newUser, 201, res);
});


exports.login = catchAsync(async (req, res, next) => {
     // 1. Check if email and password are sent
     const { email, password } = req.body;
     if(!email || !password) {
          return next(new AppError('Login credentials not sent', 400));
     }          
     // 2. Check if user exists in DB
	const user = await User.findOne({ email }).select('+password +loginAttempts +blocked' );    // we have to use 'select('+password +...')' - because we've blocked showing these users' properties when querying with find() -> we did this in the User Schema model, by using 'select: false' option on these properties
     if(!user) {
          return next(new AppError('Incorrect email or password', 404));  
     }          
     // 3. Check if user is blocked (due to excessive number of logging attempts)
     if(user.blocked) {
          // If blocked and blockExpiration time has not expired -> return error
          if(user.blockExpiration > Date.now()) {
               const timeToUnblock = (user.blockExpiration - Date.now()) / 1000;  // devide by 1000 to get seconds
               const minutes = Math.floor(timeToUnblock / 60);
               const seconds = Math.floor(timeToUnblock % 60);  // to get reminder in seconds
               return next(new AppError(`You have too many incorrect log in attempts. You are temporarely blocked from logging in. Please, wait ${minutes} minutes and ${seconds} seconds before trying to log in again.`, 401))
          }
          // If blocked, but blockExpiration time has expired -> reset user's properties and proceed
          user.blocked = false;
          user.loginAttempts = 0;
          user.blockExpiration = undefined;
          await user.save({ validateBeforeSave: false });
     }
     // 4. Check if sent password is correct (i.e. same as the one stored in DB)
     const doMatch = await bcrypt.compare(password, user.password);           		
     if(!doMatch) {
          user.loginAttempts += 1;
          // If loginAttempts are 5 or more -> block user from further logging attempts
          if(user.loginAttempts >= 5) {
               user.blocked = true;
               user.blockExpiration = Date.now() + 10*60*1000;  // block for 10 minutes
          }
          await user.save({ validateBeforeSave: false });
          return next(new AppError('Incorrect email or password', 401)); 	    
     } 
     user.loginAttempts = 0;
     await user.save({ validateBeforeSave: false });
     // 5. Create login token and send response to FE	
	sendResponseWithToken(user, 200, res);	                   
});


exports.forgotPassword = catchAsync(async (req, res, next) => {                          
     // 1. Find the user with the submitted email in DB	
     const user = await User.findOne({ email: req.body.email });
     if(!user) {
          return next(new AppError('User not found!', 404)); 
     }          
     // 2. Create random token with the crypto built-in method	
     const token = crypto.randomBytes(32).toString('hex');	 
     // 3. Hash the resetToken - in order to hide it's value (so that it cannot be just copy-pasted by potential intruder in our DB)
     const hashedToken = crypto.createHash('sha256').update(token).digest('hex');
     // 4. Apply the created token and it's expiration time to the user and save them as user's properties in DB	
     user.passwordResetToken = hashedToken;	              	  
     user.passwordResetTokenExpiration = Date.now() + 10*60*1000;                	    
     await user.save({ validateBeforeSave: false});
          // 5. Send the real token (not the hashed one) to the user's email address
     const resetUrl = `${ req.protocol }://${ req.get('host') }/users/reset-password/${ token }`;   
     const message = `Dear ${ user.name } you have requested a password reset. \nClick this ${ resetUrl } link in order to proceed!`
     try {					
          await sendEmail({
               email: user.email,
               subject: 'Reset your password at Natours.com',
               message: message
          });
          // 6. Send response data to FE	
          res.status(200).json({    
               status: 'success',	
               message: 'Reset password token sent to provided email address.'
          }); 
     } catch(err) {
          // When something goes wrong set the token to undefined
          user.passwordResetToken = undefined;	   
          user.passwordResetTokenExpiration = undefined;   
          await user.save({ validateBeforeSave: false });  
          // Set the error
          return next(new AppError('There was an error during sending the email. Please, try again.', 500));  	  
     }      		    
 });


 exports.resetPassword = catchAsync(async (req, res, next) => {            
	// 1. Find the user based on the password reset token sent as part of URL
		// Arrived token is not encrypted and the one in DB is encrypted, thus we need to encrypt the arrived token in order to compare
	const hashedToken = crypto.createHash('sha256').update(req.params.token).digest('hex');	
	const user = await User.findOne({passwordResetToken: hashedToken, passwordResetTokenExpiration: { $gt: Date.now() } })
	if(!user) {
		return next(new AppError('Token is invalid or has expired!', 400));  
	}          
	// 2. If token has not expired and user is found in DB, set the new password
	user.password = req.body.password;
	user.passwordConfirm = req.body.passwordConfirm;
	user.passwordResetToken = undefined;	   
	user.passwordResetTokenExpiration = undefined;   
	await user.save();    			
	// 3. Update the passwordChangedAt property of the user - this is done automatically in the model where we created document hooks
	// 4. Create login token and send response to FE	
	sendResponseWithToken(user, 200, res);
}); 



exports.updatePassword = catchAsync(async (req, res, next) => {            
	// 1. Find the user based on the req.user (acquired in isLogged.js middleware)	
     const user = await User.findById(req.user.id).select('+password');  
     // 2. Check if given current password is correct
     const doMatch = await bcrypt.compare(req.body.passwordCurrent, user.password); 
	if(!doMatch) {
		return next(new AppError('Incorrect password provided!', 401)); 	    
	}          
	// 3. Update the password
	user.password = req.body.password;
	user.passwordConfirm = req.body.passwordConfirm; 
     await user.save();    	
     // 4. Update the passwordChangedAt property of the user - this is done automatically in the model where we created document hook		
	// 5. Create login token and send response to FE	
	sendResponseWithToken(user, 200, res);
}); 