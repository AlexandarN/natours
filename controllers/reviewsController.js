     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	


	// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
const Review = require('../models/Review');
const handlerFactory = require('./handlerFactory');

	// 3) CONSTANTs	
	
	// 4) MIDDLEWAREs	
	// 4.1)  Add Review Middleware - this is the middleware that we made in order to be able to use handlerFactory function for 'addReview' controller
exports.setTourAndUserIDs = (req, res, next) => {	    
	// Check if tour ID is specified in the POST request (req.body) and if not use tour ID specified in the URL (req.params) -> www..../tours/id/reviews	
	if(!req.body.tour) req.body.tour = req.params.tourId;
	// User ID acquired from req.user (it is put there by the  isLogged.js middleware)
	req.body.user = req.user.id;
	next();				 
}


	// 5) CONTROLLERs	
exports.getReviews = handlerFactory.getAll(Review);

exports.addReview = handlerFactory.addDoc(Review);
		
exports.getReview = handlerFactory.getDoc(Review);
	
exports.editReview = handlerFactory.editDoc(Review);

exports.deleteReview = handlerFactory.deleteDoc(Review);
	