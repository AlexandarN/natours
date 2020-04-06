     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
	
	// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
const Tour = require('../models/Tour');
const catchAsync = require('../utils/catchAsync');
const handlerFactory = require('./handlerFactory');
	
	// 3) CONSTANTs	

	
	// 4) CONTROLLERs	
exports.getTours = handlerFactory.getAll(Tour, {path: 'reviews', select: '-__v'});

exports.addTour = handlerFactory.addDoc(Tour);

exports.getTour = handlerFactory.getDoc(Tour, {path: 'reviews', select: '-__v'});       

exports.editTour = handlerFactory.editDoc(Tour);

exports.deleteTour = handlerFactory.deleteDoc(Tour);
// 	// This is how it would look like without handler function
// exports.deleteTour = catchAsync(async (req, res, next) => {
// 	// Find a tour in DB and delete it
// 	const tour = await Tour.findByIdAndDelete(req.params.id);
// 	// In case tour is not found - send error	
// 	if(!tour) {
// 		return next(new AppError('No tour found with that ID', 404));  
// 	}          
// 	// Send response data to FE
// 	res.status(204).json({
// 		status: 'success',
// 	});
// });


exports.getTourStats = catchAsync(async (req, res, next) => {
	//  Create DB query to get the demanded statistics
	const stats = await Tour.aggregate([
		{$match: {ratingsAverage: {$gte: 4.5}}},   				       
		{$group: {
			_id: {$toUpper: '$difficulty'}, 	
			numTours: {$sum: 1},
			numRatings: {$sum: '$ratingsQuantity'},
			avgRating: {$avg: '$ratingsAverage'},					           
			avgPrice: {$avg: '$price'},
			minPrice: {$min: '$price'},
			maxPrice: {$max: '$price'} 
			}
		},
		{$sort: {numTours: -1}},
		// {$match: {_id: {$ne: 'EASY'}}}
	]);
	// Send response data to FE	
	res.status(200).json({        
		status: 'success',
		results: stats.length,
		data: {
			stats: stats } 			       
	}); 
});

exports.getMonthlyToursPlan = catchAsync(async (req, res, next) => {
	//  Create DB query to get the demanded statistics
	const year = +req.params.year;
	const plan = await Tour.aggregate([
		{$unwind: '$startDates'},
		{$match: 
			{startDates: {
					$gte: new Date(`${year}-01-01`),
					$lte: new Date(`${year}-12-31`)
				}
			}
		},
		{$group: {
				_id: {$month: '$startDates'}, 
				numStarts: {$sum: 1},
				tours: {$push: '$name'}
			}
		}, 
		{$addFields: {month: '$_id'}},
		{$project: {_id: 0}},
		{$sort: {numStarts: -1}},
		{$limit: 12}
	]);
	// Send response data to FE	
	res.status(200).json({        
		status: 'success',
		results: plan.length,
		data: {
			plan: plan } 			       
	}); 
});
