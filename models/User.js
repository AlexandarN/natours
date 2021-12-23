     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
const mongoose = require('mongoose');
const validator = require('validator');
const bcrypt = require('bcryptjs');

     // 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
const env = require('../config/env');      

     // 3) CONSTANTs


     // 4)  MODEL SCHEMA	
const userSchema = new mongoose.Schema({		              
     name: {
          type: String,				     									   
          required: [true, 'User must have a name'],                     
          trim: true,
          maxlength: [40, `User's name must have less than 40 characters'`],
          minlength: [3, `User's name must have at least 3 characters`],
          // validate: [validator.isAlphanumeric, 'Tour name must contain only numbers and letters']
     },		
     email: {
          type: String,
          required: [true, 'User must have an email'],   
          unique: true,                  
          trim: true,
          lowercase: true,
          validate: [validator.isEmail, 'Email is not valid!']
     },						             
     password: {
          type: String,
          required: [true, 'Please enter password'],
          minlength: [8, 'Password must be at least 8 characters long'],
          select: false        // we use 'select: false' to prevent displaying (cathing) password when using find() queries     	
     },						             
     passwordConfirm: {
          type: String,
          required: [true, 'Please confirm your password!'],
          validate: {
                    // NOTE: These validate functions work only on CREATE() or SAVE() and not on UPDATE()!!!	
               validator: function(val) {			   
                    return val === this.password; 
               },	           
               message: 'Passwords do not match!' 
          } 
     },
     passwordChangedAt: Date,
     photo: String,						             
     role: {
          type: String,
          enum: ['user', 'guide', 'lead-guide', 'admin'],
          default: 'user'
     },
     passwordResetToken: String,						        
     passwordResetTokenExpiration: Date,
     active: {
		type: Boolean,
		default: true,
		select: false                     
     },
     loginAttempts: {
          type: Number,
          default: 0,
          select: false
     },
     blocked: {
          type: Boolean,
          default: false,
          select: false
     },
     blockExpiration: Date
});								             
                    

     // 5) VIRTUAL PROPERTIES  	       
     
     
     // 6) MODEL METHODS	
	// 6.1) INSTANCE METHODS  --> metode koje se primenjuju nad pojedinačnim objektima Modela     

	// 6.2) STATIC METHODS (this means - this Model)	


     // 7) MIDDLEWAREs (HOOKs)
     // 7.1)  DOCUMENT MIDDLEWAREs - To be executed before or after SAVE() or CREATE(), i.e. when we are creating a resource and when we 	are actually getting an object as result of query execution 	*this - means this (new) object  
     // 7.1.1  Middleware for Hashing (encrypting) passwords 
userSchema.pre('save', async function(next) {      
          // If we are importing User data from some json file (that contains already hashed passwords) -> then do not hash passwords	  	
     if(env.NODE_ENV === 'import') {
          this.isNew = true;    // we have to set this.isNew = true; in order to avoid activation of 6.1.2 middleware (isNew is a Mongoose built-in property that checks if document is newly saved in DB)
		return next();               									          
	}
          // Only run this function if password was either modified (updated) or first time saved	
     if(!this.isModified('password')) {       
          return next(); 
     }   		         
          // Hash the password	
     this.password = await bcrypt.hash(this.password, 12);    
          // Delete the passwordConfirm field in User object	
     this.passwordConfirm = undefined;	  
     next();
});       


     // 7.1.2  Middleware for updating of passwordChangedAt property - this is activated when user is changing his password
userSchema.pre('save', async function(next) {      
          // Only run this function if password was modified (updated) and not when first time saved
     if(!this.isModified('password') || this.isNew) {       
          return next(); 
     }   		         
          // Update the passwordChangedAt property	
     this.passwordChangedAt = Date.now() - 1000;  // substract current time by 1 sec in order to avoid this time to be older than the time of issuance of jwt token (used for log in purposes, see isLogged.js middleware)    
     next();
});      


     // 7.1.3  Middleware to console.log a newly saved user
userSchema.post('save', function(doc, next) {       	  	    
     console.log(doc);    	  
     next();										
});


     // 7.2) QUERY MIDDLEWAREs - To be executed before or after FIND...(), i.e when we are updating or deleting a resource and when we are 	actually not getting any object as result of a query execution	  *this - means this query	  
     // 7.2.1  Middleware to not allow to catch (to block showing) inactive users
userSchema.pre(/^find/, function(next) {
     this.find({active: {$ne: false}});                          
     next();															
});
     
     // 7.3) AGGREGATION MIDDLEWAREs   

     // 7.4) MODEL MIDDLEWARES


     // 8) EXPORT MODEL  
module.exports = mongoose.model('User', userSchema);    
     