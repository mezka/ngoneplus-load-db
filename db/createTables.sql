DROP SCHEMA IF EXISTS Public CASCADE;
CREATE SCHEMA Public;

CREATE TABLE Public.Categories(
  categoryId SERIAL PRIMARY KEY,
  categoryName VARCHAR(100) NOT NULL
);

CREATE TABLE  Public.Products(
  productId SERIAL PRIMARY KEY,
  productName VARCHAR(100) NOT NULL,
  categoryId INTEGER REFERENCES Public.Categories(categoryId)
);

CREATE TABLE Public.Bundles(
  bundleId SERIAL PRIMARY KEY,
  bundleName VARCHAR(100) NOT NULL,
  discount FLOAT NOT NULL 
);

CREATE TABLE Public.BundleImages(
  bundleImageId SERIAL PRIMARY KEY,
  bundleId INTEGER REFERENCES Public.Bundles(bundleId) ON DELETE CASCADE,
  imageUrl VARCHAR(100)
);

CREATE TABLE Public.BundleProducts(
  bundleId INTEGER REFERENCES Public.Bundles(bundleId),
  productId INTEGER REFERENCES Public.Products(productId)
);

CREATE TABLE Public.BundleCategories(
  bundleId INTEGER REFERENCES Public.Bundles(bundleId),
  categoryId INTEGER REFERENCES Public.Categories(categoryId)
);

CREATE TABLE Public.Options(
  optionId SERIAL PRIMARY KEY,
  productId INTEGER REFERENCES Public.Products(productId) ON DELETE CASCADE,
  optionName VARCHAR(50),
  optionImageColor VARCHAR(7) DEFAULT NULL,
  optionImageIcon VARCHAR(100) DEFAULT NULL,
  optionDimensions VARCHAR(100),
  optionWeight VARCHAR(50),
  optionMaterials VARCHAR(60),
  optionPrice DOUBLE PRECISION
);

CREATE TABLE Public.OptionImages(
  imageId SERIAL PRIMARY KEY,
  optionId INTEGER REFERENCES Public.Options(optionId) ON DELETE CASCADE,
  imageUrl VARCHAR(100)
);

CREATE TABLE Public.Users(
  userId SERIAL PRIMARY KEY,
  stripeId VARCHAR(50) NOT NULL UNIQUE,
  userEmail VARCHAR(60) NOT NULL,
  userFirstName VARCHAR(50) NOT NULL,
  userLastName VARCHAR(50) NOT NULL
);

CREATE TABLE Public.Address(
  id SERIAL PRIMARY KEY,
  userId INTEGER REFERENCES Public.Users(userId) ON DELETE CASCADE,
  address1 VARCHAR(50) NOT NULL,
  address2 VARCHAR(50),
  city  VARCHAR(50),
  state VARCHAR(50),
  country VARCHAR(50) NOT NULL,
  zipcode  VARCHAR(50) NOT NULL
);

CREATE TABLE Public.LocationData(
  id SERIAL PRIMARY KEY,
  address_id INTEGER REFERENCES Public.Address(id) ON DELETE CASCADE,
  country_iso_code VARCHAR(2) NOT NULL,
  state_id VARCHAR(2),
  geoname_id VARCHAR(8)
);

CREATE TABLE Public.Order(
  id SERIAL PRIMARY KEY,
  userId INTEGER REFERENCES Public.Users(userId) ON DELETE CASCADE,
  addressId INTEGER REFERENCES Public.Address(id) NOT NULL,
  createdDate timestamptz NOT NULL DEFAULT now(),
  paid BOOLEAN DEFAULT FALSE,
  amount DOUBLE PRECISION DEFAULT NULL,
  receipt_url VARCHAR(140) DEFAULT NULL
);

CREATE TABLE Public.OrderItem(
  id SERIAL PRIMARY KEY,
  orderId INTEGER REFERENCES Public.Order(id) ON DELETE CASCADE NOT NULL,
  productId INTEGER REFERENCES Public.Products(productId) ON DELETE CASCADE NOT NULL,
  optionId INTEGER REFERENCES Public.Options(optionId) ON DELETE CASCADE NOT NULL,
  quantity INTEGER NOT NULL,
  price DOUBLE PRECISION,
  discount INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE Public.Password(
    id SERIAL PRIMARY KEY,
    userId INTEGER REFERENCES Public.Users(userId) ON DELETE CASCADE,
    passwordHash VARCHAR(60) UNIQUE
)
