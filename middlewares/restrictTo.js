     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	

     // 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)
const AppError = require('../utils/AppError');

	// 3) CONSTANTs	

     // 4) MIDDLEWAREs	
module.exports = (...roles) => {             // Acquire allowed roles -> it is an array of values
     return (req, res, next) => {
          // Check if the logged in user's role is part of the allowed roles array
          if(!roles.includes(req.user.role)) {
               return next(new AppError('You do not have permission to perform this action', 403));
          }
          // If user's role is allowed (is part of the roles array) proceed
          next();
     }
}