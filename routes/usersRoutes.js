     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
const express = require('express');


	// 2) IMPORTING our custom files (CONTROLLERs, MIDDLEWAREs, ...)	
const usersController = require('../controllers/usersController');
const authController = require('../controllers/authController');
const isLogged = require('../middlewares/isLogged');
const restrictTo = require('../middlewares/restrictTo');

	// 3) CONSTANTs	
const router = express.Router();


	// 4.1) AUTHENTICATION ROUTES
router.post('/signup', authController.signup);
router.post('/login', authController.login);
router.post('/forgot-password', authController.forgotPassword);
router.patch('/reset-password/:token', authController.resetPassword);


	//  MIDDLEWARE	
router.use(isLogged);		// From here to below all routes request the use of 'isLogged' middleware 

	// 4.1) AUTHENTICATION ROUTES
router.patch('/update-password', authController.updatePassword); // user updates his own password when already logged in


	// 4.2)  SELF-UPDATE ROUTES	
router.get('/get-me', usersController.getMe, usersController.getUser);  // 'getMe' is a midd. we created in order to be able to use 'getDoc' handler
router.patch('/update-me', usersController.updateMe);	// user updates his own data		
router.delete('/delete-me', usersController.deleteMe);	


	//  MIDDLEWARE	
router.use(restrictTo('admin'));		// From here to below all routes request also the use of 'restrictTo' middleware			              

	// 4.3) CRUD ROUTEs	
router.route('/')
	.get(usersController.getUsers)
	.post(usersController.addUser);

router.route('/:id')
	.get(usersController.getUser)
	.patch(usersController.editUser)
	.delete(usersController.deleteUser);
	

	// 5) EXPORT ROUTER	
module.exports = router;