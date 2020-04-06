class QueryFeatures {
     constructor(modelQuery, reqQuery) {
          this.modelQuery = modelQuery;
          this.reqQuery = reqQuery;
     }
   
     // FILTERING
     filter() {
          // 1A - FILTERING resources based on parameters sent in req.query
          const queryObj = { ...this.reqQuery };
          const excludedFields = ['sort', 'fields', 'page', 'limit'];
          excludedFields.forEach(el => delete queryObj[el]);
   
          // 1B - ADVANCED FILTERING when having MongoDB Query OPERATORS sent as parameters in req.query e.g: .../tours?duration[gte]=5
          let filterStr = JSON.stringify(queryObj);
          filterStr = filterStr.replace(/\b(gte|gt|lte|lt)\b/g, match => `$${match}`);

               // Create 1st query based on sent filtering parameters
          this.query = this.modelQuery.find(JSON.parse(filterStr));
          console.log(JSON.parse(filterStr));
          return this;
     }
   
     // SORTING
     // 2 - ADVANCED SORTING based on 2 or more parameters sent in req.query www.../tours?sort=price,duration
     sort() {
          if (this.reqQuery.sort) {
               console.log(this.reqQuery.sort);
               const sortStr = this.reqQuery.sort.split(',').join(' ');
               // Create 2nd query based on sent sorting parameters
               this.query = this.modelQuery.sort(sortStr);
          } else {
               // Create 2nd query based on default sorting parameters
               this.query = this.modelQuery.sort('-createdAt');
          }
          return this;
     }
   
     // PROJECTING
	// 3 - PROJECTING which FIELDS (properties) to be caught based on params sent in req.query separated by coma /tours?fields=name,price
     project() {
          if (this.reqQuery.fields) {
               const fieldStr = this.reqQuery.fields.split(',').join(' ');
               // Create 3rd query based on sent fields parameters
               this.query = this.modelQuery.select(fieldStr);
          } else {
               // Create 3rd query based on default fields parameters
               this.query = this.modelQuery.select('-__v');
          }
   
          return this;
     }
   
     // PAGINATION	
     paginate() {
          // 4A - SKIPPING resources based on page parameter sent in req.query
          const page = +this.reqQuery.page || 1;
          // 4B - LIMITTING resources based on limit parameter sent in req.query
          const limit = +this.reqQuery.limit || 70;
          // Create 4th query based on sent or default page and limit parameters
          this.query = this.modelQuery.skip((page - 1) * limit).limit(limit);
          return this;
     }
}

module.exports = QueryFeatures;