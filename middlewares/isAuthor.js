     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	

     // 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)
const AppError = require('../utils/AppError');
const Review = require('../models/Review');
const catchAsync = require('../utils/catchAsync.js'); 

	// 3) CONSTANTs	

     // 4) MIDDLEWAREs	
module.exports = catchAsync(async (req, res, next) => {
     // If logged in user's role is 'user' -> then check if he is author of the review he wants to access
     if(req.user.role === 'user') {
          const review = await Review.findById(req.params.id);
          if(review.user.id !== req.user.id) {
               return next(new AppError('You are not the author of this review and you do not have permission to perform this action', 403));
          }
     }
     // If a user is the author of the review -> proceed
     next();
});
