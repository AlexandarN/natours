module.exports = {
     NODE_ENV: process.env.NODE_ENV,
     PORT: process.env.PORT,
     // STRIPE_PUB_KEY: process.env.STRIPE_PUBLISHABLE_KEY,
     // STRIPE_SECRET_KEY: process.env.STRIPE_SECRET_KEY,
     MONGODB_USER: process.env.MONGODB_USER,
     MONGODB_PASSWORD: process.env.MONGODB_PASSWORD,
     MONGODB_DEFAULT_DB: process.env.MONGODB_DEFAULT_DB,
     JWT_SECRET: process.env.JWT_SECRET,
     JWT_EXPIRATION: process.env.JWT_EXPIRATION,
     JWT_COOKIE_EXPIRATION: process.env.JWT_COOKIE_EXPIRATION,
     JWT_COOKIE_SECURE: true,
     EMAIL_USERNAME: process.env.EMAIL_USERNAME,
     EMAIL_PASSWORD: process.env.EMAIL_PASSWORD,
     EMAIL_HOST: process.env.EMAIL_HOST,
     EMAIL_PORT: process.env.EMAIL_PORT
}