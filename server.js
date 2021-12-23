/* eslint-disable no-console */
/* eslint-disable import/newline-after-import */
     // 1) IMPORTING NPM PACKAGEs and NODE MODULEs	
const mongoose = require('mongoose');	
     // To avoid some deprication warnings                                                    
mongoose.set('useCreateIndex', true);   
mongoose.set('useFindAndModify', false);


     // 7.1)  Catching UNCAUGHT EXCEPTION ERRORS (that are result of some strange unexpected error) -> we need to put this above all the code
// process.on('uncaughtException', err => {
//      console.log('UNCAUGHT EXCEPTION!');
//      console.log('ERROR: ', err.name, err.message);
// });

     // 2) IMPORTING our custom files (APP.js and ENV.js)
const app = require('./app');
const env = require('./config/env');      			


	// 3) CONSTANTs								    
// eslint-disable-next-line prefer-template
const MONGODB_URI  = 'mongodb+srv://' + env.MONGODB_USER + ':' + env.MONGODB_PASSWORD + '@cluster0-eyxah.mongodb.net/' + env.MONGODB_DEFAULT_DB;
// const MONGODB_URI  = 'mongodb://localhost/natours'; 


     // CHECKING the type of ENVIRONMENT								    
console.log(app.get('env'));
// console.log(process.env);


	// 6.1) DB CONNECTION to APP. SERVER			 
mongoose.connect(MONGODB_URI, {useNewUrlParser: true});       	       
const db = mongoose.connection;							    
db.on('error', console.error.bind(console, 'connection error:'));                            
db.once('open', () => {
     console.log('Connected to DB'); 
});


     // 6.2) STARTING the APP. SERVER	
app.listen(env.PORT , () => {
     console.log(`Server started on port ${env.PORT}`);
});


	// 7.2) Catching UNHANDLED REJECTION ERRORS (that aren't result of our code, but are coming from the outside, e.g. rejected connection to DB)      
process.on('unhandledRejection', err => {
     console.log('UNHANDLED REJECTION!');
     console.log('ERROR: ', err.name, err.message);
});


     