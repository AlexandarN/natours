     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
const express = require('express');


	// 2) IMPORTING our custom files (CONTROLLERs, MIDDLEWAREs, ...)	
const reviewsController = require('../controllers/reviewsController');
const isLogged = require('../middlewares/isLogged');
const restrictTo = require('../middlewares/restrictTo');
const isAuthor = require('../middlewares/isAuthor');

	// 3) CONSTANTs	
const router = express.Router({ mergeParams: true});   // *sa {mergeParams: true} omogućavamo da se u ovom Routeru prihvataju poslati parametri u req.params koji se odnose na druge modele (tj. npr. tourId kao ID od Tour modela), tako da ako imamo rutu www..../tours/tourId/reviews ovaj Router će onda dohvatiti tourId. Ovakva ruta će se ovde u review Routeru tretirati isto kao i www..../reviews , te će se aktivirati isti kontroleri u oba slučaja. 


	//  MIDDLEWARE	
router.use(isLogged);		// From here to below all routes request the use of 'isLogged' middleware 

	// 4) CRUD ROUTEs	
router.route('/')
	.get(reviewsController.getReviews)
	.post(restrictTo('user'), reviewsController.setTourAndUserIDs, reviewsController.addReview);

router.route('/:id')
	.get(reviewsController.getReview)
	.patch(restrictTo('user', 'admin'), isAuthor, reviewsController.editReview)
	.delete(restrictTo('user', 'admin'), isAuthor, reviewsController.deleteReview);
	

	// 5) EXPORT ROUTER	
module.exports = router;