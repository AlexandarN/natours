     // 1)  IMPORTING NPM PACKAGEs and NODE MODULEs	


     // 2)  IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	
const AppError = require('../utils/AppError');   			 
const catchAsync = require('../utils/catchAsync');   		
const QueryFeatures = require('../utils/QueryFeatures');  

     // 3)  CONSTANTs	
     
     
     // 4)  HANDLERs
exports.getAll = (Model, popOptions) => catchAsync(async (req, res, next) => {    
     // Only for Get All Reviews controller we need to check for tourIdFilter - i.e. Check if tour ID is specified in URL (req.params) -> www..../tours/id/reviews
	let tourIdFilter = {};	
	if(req.params.tourId) {
		tourIdFilter = { tour: req.params.tourId };  	// if tour ID is specified in req.params then find reviews that belong to the specified tour
     }
     // Check if popOptions are given
     let modelQuery = Model.find(tourIdFilter);
     if(popOptions) {
          modelQuery = modelQuery.populate(popOptions);        
     }
     // EXECUTE QUERY - Find the requested documents in DB based on the created query
     const featuresObj = new QueryFeatures(modelQuery, req.query).filter().sort().project().paginate();
     // const docs = await featuresObj.query.explain();  ---> we use explain() when we want to get statictics about the efiiciency of execution of this query
     const docs = await featuresObj.query;  
     // Send response data to FE	      
     res.status(200).json({        	                                				  	  
          status: 'success',       	       
          results: docs.length,
          data: {
               documents: docs 
          } 				     	 
     }); 
});
     

exports.addDoc = Model => catchAsync(async (req, res, next) => {
     // Create new document in DB	
     const doc = await Model.create(req.body);           
     // Send response data to FE	
     res.status(201).json({    
          status: 'success',		                 					      
          data: {
               document: doc } 				          	                   
     }); 
});


exports.getDoc = (Model, popOptions) => catchAsync(async (req, res, next) => {
     // Check if popOptions are given
     let query = Model.findById(req.params.id);
     if(popOptions) {
          query = query.populate(popOptions);        
     }
     // EXECUTE QUERY - Find the requested document in DB	
     const doc = await query;  
     // In case document is not found - send error	
     if(!doc) {
          return next(new AppError('No document found with that ID', 404));
     }
     // Send response data to FE											       
     res.status(200).json({    
          status: 'success',		                 					      
          data: {
               document: doc } 				          	          	          
     }); 
});  


exports.editDoc = Model => catchAsync(async (req, res, next) => {        
     // Find a document in DB and update it		   *Napomena: - ovo će nam raditi samo ako je req poslat na route sa patch metodom
     const doc = await Model.findByIdAndUpdate(req.params.id, req.body, {  	 
          new: true,				 	  
          runValidators: true     
     });
     // In case document is not found - send error	
     if(!doc) {
          return next(new AppError('No document found with that ID', 404));
     }          
     // Send response data to FE	
     res.status(200).json({    
          status: 'success',		                 					      
          data: {
               document: doc 
          } 				          	
     }); 
});
     
exports.deleteDoc = Model => catchAsync(async (req, res, next) => {
	// Find a document in DB and delete it
	const doc = await Model.findByIdAndDelete(req.params.id);
	// In case document is not found - send error	
	if(!doc) {
		return next(new AppError('No document found with that ID', 404));  
	}          
	// Send response data to FE
	res.status(204).json({
		status: 'success',
	});
});