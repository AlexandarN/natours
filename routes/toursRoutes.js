/* eslint-disable import/newline-after-import */
     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
const express = require('express');


	// 2) IMPORTING our custom files (CONTROLLERs, MIDDLEWAREs, ...)	
const toursController = require('../controllers/toursController');
const reviewsRoutes = require('../routes/reviewsRoutes');

const aliasTours = require('../middlewares/aliasTours');
const isLogged = require('../middlewares/isLogged');
const restrictTo = require('../middlewares/restrictTo');


	// 3) CONSTANTs	
const router = express.Router();


	// CATCHING HTTP Request Parameters with ROUTER
router.param('id', (req, res, next, val) => {
	console.log(`Tour id is: ${val}`);
	next();
});


	// 4.1) ALIAS ROUTEs	
router.route('/top5tours')
	.get(aliasTours.top5Tours, toursController.getTours);
router.route('/tour-stats')
	.get(toursController.getTourStats);	
router.route('/monthly-plan/:year')
	.get(isLogged, restrictTo('admin', 'lead-guide', 'guide'), toursController.getMonthlyToursPlan);


	// 4.2)  GEOSPATIAL ROUTEs	
router.route('/tours-within/:distance/center/:coordinates/unit/:unit', toursController.getToursWithin)


	// 4.3) CRUD ROUTEs	
router.route('/')
	.get(toursController.getTours)
	.post(isLogged, restrictTo('admin', 'lead-guide'), toursController.addTour);

router.route('/:id')
	.get(toursController.getTour)
	.patch(isLogged, restrictTo('admin', 'lead-guide'), toursController.editTour)
	.delete(isLogged, restrictTo('admin', 'lead-guide'), toursController.deleteTour);


	// 4.4)  NESTED ROUTEs -> for PARENT - CHILD relationship (routes that start with '/tours' but will be handled by other model's router)	
router.use('/:tourId/reviews', reviewsRoutes);			// for this exact route use reviewsRoutes instead of tourRoutes
// router.route('/:tourId/reviews')
// 	.get(reviewsController.getReviews)
// 	.post(isLogged, restrictTo('user'), reviewsController.addReview);
	

	// 5) EXPORT ROUTER	
module.exports = router;

