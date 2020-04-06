     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	

	// 2) IMPORTING our custom files (MODELs, MIDDLEWAREs, ...)	

	// 3) CONSTANTs	

     // 4) MIDDLEWAREs	
exports.top5Tours = (req, res, next) => {
     req.query.sort = '-ratingsAverage,price';
     req.query.fields = 'name,ratingsAverage,price,difficulty,summary,duration';
     req.query.limit = 5;
     next();
}