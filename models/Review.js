     // 1)  IMPORTING NPM PACKAGEs and NODE MODULEs	
const mongoose = require('mongoose');


	// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
const Tour = require('../models/Tour');   // we need this in the static method 6.2.1

	// 3) CONSTANTs

	// 4)  MODEL SCHEMA	
const reviewSchema = new mongoose.Schema({		             
	review: {
		type: String,	
		required: [true, 'Review must have content'],                  
		maxlength: [400, 'Review must have less than 400 characters'],
		minlength: [3, 'Review must have at least 3 characters']	      	       
		},				         
	rating: {
          type: Number,
		min: [1, 'Rating must be minimum 1.0'],
		max: [5, 'Rating must be maximum 5.0']       			    
	},						             
	createdAt: {
		type: Date,
		default: Date.now()
	},
	user: {
          type: mongoose.Schema.ObjectId,   		        	     
          ref: 'User',
          required: [true, 'Review must have an author!']
	},
	tour: {
          type: mongoose.Schema.ObjectId,   		       	       
          ref: 'Tour',
          required: [true, 'Review must belong to a tour!']
	}
},
{   
	toJSON: {virtuals: true},   
	toObject: {virtuals: true}			        
}
);		

    	//  4.2) MODEL INDEXING 
reviewSchema.index({ tour: 1, user: 1 }, { unique: true });    // {unique: true} - option that sets that there can be only one review with the same tour and same user combination pair (i.e. a user can post only one review on the same tour, he is forbidden to post more reviews on the same tour)   

	
	// 5) VIRTUAL PROPERTIES


	// 6) MODEL METHODS	
	// 6.1) INSTANCE METHODS  --> metode koje se primenjuju nad pojedinačnim objektima Modela     


	// 6.2)  STATIC METHODS (this means - this Model)	
	// 6.2.1)  Method for calculation of 'averageRatings' -> which is also property of Tour objects (it needs to be calculated and updated each time new review is saved, i.e. when someone posts his rating on a specific tour)	
reviewSchema.statics.calculateAvgRatings = async function(tourId) {
		// After review is saved in DB, get all reviews of the same tour and calculate its' total no. of ratings and average rating
	const stats = await this.aggregate([	         					       
		{$match: {tour: tourId}},
		{$group: {  
			_id: '$tour',   	      
			numRatings: {$sum: 1},	  
			avgRating: {$avg: '$rating'} 
			}	                               
		}
	]);
	console.log(stats);
		// If there are any reviews on the given tour there will be some results in the 'stats' variable (otherwise 'stats' variable will be an empty object) -> so, if there are any reviews find the tour in DB and update its' properties
	if(stats.length > 0) {
		await Tour.findByIdAndUpdate(tourId, {
			ratingsQuantity: stats[0].numRatings,    
			ratingsAverage: stats[0].avgRating          
		});
	} else {
		// If there are not any reviews on the given tour ('stats' variable is empty) -> this could happen when deleting all reviews of a tour
		await Tour.findByIdAndUpdate(tourId, {
			ratingsQuantity: 0,
			ratingsAverage: 0							      
		});
	}
}			


	// 7) MIDDLEWAREs (HOOKs)
	// 7.1)  DOCUMENT MIDDLEWAREs - To be executed before or after SAVE() or CREATE(), i.e. when we are creating a resource and when we 	are actually getting an object as result of query execution 	*this - means this (new) object
	// 7.1.1  Middleware for calling (executing) 'calculateAvgRatings' STATIC method (6.2.1) when we are creating new review
reviewSchema.post('save', function() {          
	this.constructor.calculateAvgRatings(this.tour);      
});
			

	// 7.2)  QUERY MIDDLEWAREs - To be executed before or after FIND...(), i.e when we are updating or deleting a resource and when we are 	actually not getting any object as result of a query execution	  *this - means this query
	// 7.2.1  When we have REFERENCING relationship between 2 MODELs we use this Query Middleware to always catch properties of the referenced model object (here we catch users by using populate() Mongoose method on Review.find() )
reviewSchema.pre(/^find/, function(next) {            
	this.populate({	         					      
		path: 'user', 					   
		select: 'name photo'				
	});
	next();															
});

 	// 7.2.2  Middleware to catch an object that is about to be updated or deleted with findByIdAnd...() -> we need this to calculate ratings average, i.e. to be able to use 'calculateAvgRatings' static method (6.2.1) when we are updating or deleting reviews
// reviewSchema.pre(/^findOneAnd/, async function(next) {        	      
//      this.review = await this.findOne();      
//      next();    	  	  
// });
     // 7.2.3  Middleware for calling (executing) 'calculateAvgRatings' STATIC method (6.2.1) when we are creating new review
reviewSchema.post(/^findOneAnd/, async function(doc) {        	     
     await doc.constructor.calculateAvgRatings(doc.tour);       
});  


	// 8)  EXPORT MODEL	
module.exports = mongoose.model('Review', reviewSchema);    
