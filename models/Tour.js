     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
const mongoose = require('mongoose');
const slugify = require('slugify');


	// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	

	// 3) CONSTANTs


	// 4)  MODEL SCHEMA	
const tourSchema = new mongoose.Schema({		              
     name: {
          type: String,				     									   
          required: [true, 'A tour must have a name'],                     
          unique: true,
          trim: true,
          maxlength: [40, 'A tour name must have less than 41 characters'],
          minlength: [5, 'A tour name must have at least 5 characters']
          // validate: [validator.isAlphanumeric, 'Tour name must contain only numbers and letters']
     },		
     slug: String,
     duration: {
          type: Number,
          required: [true, 'A tour must have a duration']                     
     },						             
     maxGroupSize: {
          type: Number,
          required: [true, 'A tour must have a group size']                     
     },						             
     difficulty: {
          type: String,
          required: [true, 'A tour must have a difficulty level'],
          enum: {
			values: ['easy', 'medium', 'difficult'],
			message: 'Difficulty is either: easy, medium or difficult!'	      
		}			             				       
     },						             
     ratingsAverage: {
          type: Number,				              					           
          default: 0,
          min: [1, 'Rating must be minimum 1.0'],
          max: [5, 'Rating must be maximum 5.0'],
          set: val => Math.round(val * 10) / 10    // with this expression we make sure that the value of this property is always shown as a decimal number with 1 decimal (e.g. 4.7)
     },
     ratingsQuantity: {
          type: Number,				              					           
          default: 0 
     },
     price: {
          type: Number,				     									
          required: [true, 'A tour must have a price'] 
     },
     priceDiscount: {
		type: Number,
		validate: {  
               // NOTE: These validate functions work only on CREATE() or SAVE() and not on UPDATE()!!!	
			validator: function(val) {			       
                    return val < this.price; 
               },						            
               message: 'Discount price {VALUE} should be below regular price' 
          } 
	},
     summary: {
          type: String,
          required: [true, 'A tour must have a summary'],
          trim: true
     },
     description: {
          type: String,
          trim: true
     },
     imageCover: {
          type: String,
          required: [true, 'A tour must have a cover image']
     },
     images: [String],
     createdAt: {
          type: Date,
          default: Date.now(),
          select: false
     },
     startDates: [Date],
     secretTour: {
          type:Boolean,
          default: false
     },
     startLocation: {
		// GeoJSON object (must have type and coordinates)                    
		type: {			      				            
			type: String,     					
			default: 'Point',
               enum: ['Point'] 
          },			       
		coordinates: [Number],  // array of numbers -> longitude + latitude      
		address: String,
		description: String   
	},
	locations: [   								
          // Array of embedded GeoJSON objects
		{
			type: {			      			            
				type: String,     					
				default: 'Point',
                    enum: ['Point'] 
               },					      
			coordinates: [Number],   						             
			address: String,
			description: String,
			day: Number
		}
     ],
     guides: [
		{
          	type: mongoose.Schema.Types.ObjectId,   		 
               ref: 'User'
		}
	]
},
{
     toJSON: {virtuals: true}, 
     toObject: {virtuals: true}			      
}
);					

    	//  4.2) MODEL INDEXING 
tourSchema.index({price: 1, ratingsAverage: -1}); 
tourSchema.index({ slug: 1 });                              
// NOTE: All tours in DB will be indexed on 'price' property in ascending order. This practically means that they will be searched in the order starting from the lowest price to the highest. No matter what we are querying for (e.g. catch the tours with duration lower than 5 days), Mongoose will always search documents (tours) in the above stated order. The point of indexing is to fasten the search when we are searching documents by the indexed property, i.e. to stop searching after all conditions are met (e.g. after all tours with price lower than 1000 are caught). Without setting up an index on some property, all queries will be executed on all documents in the collection. The point of indexing is to set up an index on a property that is highly queried (most queried), as in this case -> searching might take shorter time (i.e. it might not be executed on all documents in collection).
                  

     // 5) VIRTUAL PROPERTIES  
     // 5.1)  DurationWeeks virtual property  	       
tourSchema.virtual('durationWeeks').get(function() {        	      
     return this.duration / 7;						      	      
});		

    // 5.2)  VIRTUAL POPULATE -> to get REVIEWS per each tour (as reviews are not part of the Tour schema, instead we use virtual population in order to get objects of the other model (childs) that belong to the object of this model (parent), since we have here PARENT REFERENCING (IDs of this model objects are part of the other model's objects (child)) -> we use VIRTUAL POPULATE when we don't want to use CHILD REFERENCING (in order to avoid having big number of review IDs stored into one tour object), as virtual properties are not actually stored in DB, instead they are calculated when we call them in controllers by using populate('reviews')
tourSchema.virtual('reviews', {    // name of virtual property in this model
     ref: 'Review',                // reference Model
     foreignField: 'tour',         // name of property in the other model
     localField: '_id'             // name of property in this model
});


	// 6) MODEL METHODS	
	// 6.1) INSTANCE METHODS  --> metode koje se primenjuju nad pojedinačnim objektima Modela     

	// 6.2) STATIC METHODS (this means - this Model)	


     // 7) MIDDLEWAREs (HOOKs)
     // 7.1)  DOCUMENT MIDDLEWAREs - To be executed before or after SAVE() or CREATE(), i.e. when we are creating a resource and when we 	are actually getting an object as result of query execution 	*this - means this (new) object	
     // 7.1.1  Middleware for automatic creation of slug for each tour when being saved    
tourSchema.pre('save', function(next) {       
     this.slug = slugify(this.name, {lower: true});         
     next();
});


     // 7.1.2 Middleware for AUTOMATIC EMBEDING OF other MODEL'S OBJECTs (based on their provided ID) into this MODEL OBJECT -> (we use this in case of EMBEDDING a User object into a Tour object, but in this project we will instead use REFERENCING (saving IDs of childs in the parent), because if there is a change in guides (users) we would have to make the change in both User object and Tour object)
// tourSchema.pre('save', async function(next) {         
//      const guidesPromises = this.guides.map(id => {
// 		return User.findById(id);	
// 	});   
// 	this.guides = await Promise.all(guidesPromises);    // The .map() function cannot actually be made asynchronous. It will always be synchronous. The end result is that guidesPromises is an array of pending promises (therefore being empty). We need to wait until all of those promises have been resolved though so we call Promise.all(guidesPromises) which does wait until every promise has been resolved before continuing.
//      next();										        
// });
          

     // 7.1.3  Middleware for printing a tour in console.log() after being saved
tourSchema.post('save', function(doc, next) {       	  	   
	console.log(doc);    	 
	next();										      
});


     // 7.2)  QUERY MIDDLEWAREs - To be executed before or after FIND...(), i.e when we are updating or deleting a resource and when we are 	actually not getting any object as result of a query execution	  *this - means this query 	
     // 7.2.1  When we have REFERENCING relationship between 2 MODELs we use this Query Middleware to always catch properties of the referenced model object (here we catch users (guides) by using populate() Mongoose method on Tour.find() )
tourSchema.pre(/^find/, function(next) {            
    this.populate({	         			// this - represents the query (Tour.find() )		       
         path: 'guides', 				 // in 'path:' we put referenced property
         select: '-passwordChangedAt  -__v'	 // in 'select' we put properties which we do not want to be shown (caught) by the query     
    });
    next();															
});

     // 7.2.2  Middleware to not allow to catch (to block showing) tours with secretTour = true     
tourSchema.pre(/^find/, function(next) {             
     this.find({secretTour: {$ne: true}});   
     this.queryStartTime = Date.now();
	next();										
});

     // 7.2.3  Middleware for calculation of time needed to execute each find tours query
tourSchema.post(/^find/, function(docs, next) {                  
	console.log(`Mongoose Middleware: This query execution lasts ${(Date.now() - this.queryStartTime) / 1000} seconds.`);  
	next();										        
});


     // 7.3) AGGREGATION MIDDLEWAREs  
     // 7.3.1  Middleware for changing of aggregation pipeline query 
tourSchema.pre('aggregate', function(next) {                
     this.pipeline().unshift({$match: {duration: {$gt: 0}}});  // here with unshift() - we add another query {$match: ...} to the aggragation pipeline
     console.log(this.pipeline());    	         
     next();													      
});

     // 7.4) MODEL MIDDLEWAREs - these also exist but are not so important


     // 8) EXPORT MODEL
module.exports = mongoose.model('Tour', tourSchema);    
