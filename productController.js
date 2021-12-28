const tmp = require('tmp');
const util = require('util');
const fs = require('fs');
const exceljs = require('exceljs');
const _ = require('lodash');
const { v4: uuidv4 } = require('uuid');
const { ObjectId } = require('mongoose').Types;
const moment = require('moment');

const environments = require('../../config/environments');
const error = require('../../middlewares/errorHandling/errorConstants');
const { Product, ModifiedProduct, SoonInStock, Store, brandTypes, collectionTypes, Client, statuses, Activity, rateRSDWatches, rateHUFWatches, vatRs, vatHu, vatMn, watchLocationsBG, watchLocationsBU, watchLocationsPM, multibrandLocations, Wishlist, Report, Checkout, jewelryMaterials, rateRSDJewelry, Shipment, jewelryTypes, stoneTypes, colors, clarities, shapes, cuts, LabelSerial, onStatuses, Shortlist } = require('../../models');
const { minioClient, deleteFile } = require('../../lib/fileHandler');
const { createActivity, isValidId, setBraceletType, setCaseMaterial, setRolexMaterialsAndCollections, getProductGroup, getRecentPurchaseValidation } = require('../../lib/misc');
const { stockCheck, labelsPDF } = require('../../lib/pdfHandler');
const { getProductLabelText, getProductDeclaration, getProductQRDataURL, transformSoonInStockToProductDoc } = require('./productHelper');

/**
 * @api {post} /product Import products
 * @apiVersion 1.0.0
 * @apiName importProductModels
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (formData) {Binary} file Binary file uploaded
 * @apiParam (query) {String=true} [save] Required when confirming the review (i.e. saving changes)
 * @apiParam (query) {String='Rolex', 'Tudor', 'Panerai', 'SwissKubik', 'Rubber B'} brand Brand Type
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 201 OK
 {
   "message": "Successfully imported new product models",
 }
 *
 * @apiUse MissingParamsError
 * @apiUse InvalidValue
 * @apiUse InvalidColumnName
 * @apiUse InvalidMaterial
 * @apiUse MissingPrice
 * @apiUse MissingPurchasePrice
 * @apiUse CredentialsError
 */
module.exports.importProductModels = async (req, res) => {
  const { objectName } = req.file;
  const { save, brand } = req.query;

  // Check if req.file has been sent
  if (!objectName || !brand) throw new Error(error.MISSING_PARAMETERS);
  if (!['Rolex', 'Tudor', 'Panerai', 'SwissKubik', 'Rubber B'].includes(brand)) throw new Error(error.INVALID_VALUE);

  // Create a stream to pull excel file with particular 'objectName' from the Minio server
  const dataStream = await minioClient.getObject(environments.MINIO_BUCKET, objectName);

  // Create an empty workbook
  const workbook = new exceljs.Workbook();

  // Fill in the workbook with the excel data pulled from the Minio server
  await workbook.xlsx.read(dataStream);

  // Set 'worksheet' variable -> that contains data from the 1st sheet of 'workbook' (uploaded excel file)
  const worksheet = workbook.getWorksheet(1);

  // Get existing stores in DB
  const existingStores = await Store.find().lean();

  // Get existing products in DB with the specified 'brand'
  const existingProducts = await Product.find({ brand }).lean();

  // I - Insert products data from the uploaded 'RMC' file into DB
  let newProducts = [];
  let updatedProducts = [];
  let modifiedProducts = [];
  const rmcArray = [];

  // 1. Delete all modified products in DB, that have been created during previous upload of products
  await ModifiedProduct.deleteMany({});

  // 2. After confirming (saving) the table, set the status of all existing products in DB, whose 'status' isn't 'deleted' -> to 'previous'
  if (save) await Product.updateMany(
    { status: { $ne: 'deleted' }, brand },
    { status: 'previous' }
  ).lean();

  // Set photo extension (if Rolex brand)
  let extension = 'jpg';

  switch (brand) {
    case 'Tudor':
    case 'Panerai':
    case 'SwissKubik':
    case 'Rubber B':
      extension = 'png';
      break;
  }

  // Create new product model for each row in the RMC table
  for (let i = 0; i < worksheet.actualRowCount - 1; i += 1) {
    switch (brand) {
      case 'Rolex':
        // Check columns order
        if (
          worksheet.getCell(`A1`).value !== 'RMC' ||
          worksheet.getCell(`B1`).value !== 'Collection' ||
          worksheet.getCell(`C1`).value !== 'Product Line' ||
          worksheet.getCell(`D1`).value !== 'Sale Reference' ||
          worksheet.getCell(`E1`).value !== 'Material Description' ||
          worksheet.getCell(`F1`).value !== 'Dial' ||
          worksheet.getCell(`G1`).value !== 'Bracelet' ||
          worksheet.getCell(`H1`).value !== 'Box' ||
          worksheet.getCell(`I1`).value !== 'Diameter' ||
          worksheet.getCell(`J1`).value !== 'Case Material' ||
          worksheet.getCell(`K1`).value !== 'Ex_Geneve_CHF' ||
          worksheet.getCell(`L1`).value !== 'Retail_EUR_RS' ||
          worksheet.getCell(`M1`).value !== 'Retail_EUR_HU' ||
          worksheet.getCell(`N1`).value !== 'Retail_EUR_MNE'
        ) throw new Error(error.INVALID_COLUMN_NAME);
        break;

        case 'Tudor':
        // Check columns order
        if (
          worksheet.getCell(`A1`).value !== 'RMC' ||
          worksheet.getCell(`B1`).value !== 'Collection' ||
          worksheet.getCell(`C1`).value !== 'Product Line' ||
          worksheet.getCell(`D1`).value !== 'Sale Reference' ||
          worksheet.getCell(`E1`).value !== 'Material Description' ||
          worksheet.getCell(`F1`).value !== 'Dial' ||
          worksheet.getCell(`G1`).value !== 'Bracelet' ||
          worksheet.getCell(`H1`).value !== 'Box' ||
          worksheet.getCell(`I1`).value !== 'Diameter' ||
          worksheet.getCell(`J1`).value !== 'Case Material' ||
          worksheet.getCell(`K1`).value !== 'Waterproofness' ||
          worksheet.getCell(`L1`).value !== 'Ex_Geneve_CHF' ||
          worksheet.getCell(`M1`).value !== 'Retail_EUR_RS'
        ) throw new Error(error.INVALID_COLUMN_NAME);
        break;

      case 'Panerai':
        // Check columns order
        if (
          worksheet.getCell(`A1`).value !== 'RMC' ||
          worksheet.getCell(`B1`).value !== 'Collection' ||
          worksheet.getCell(`C1`).value !== 'Sale Reference' ||
          worksheet.getCell(`D1`).value !== 'Material Description' ||
          worksheet.getCell(`E1`).value !== 'Dial' ||
          worksheet.getCell(`F1`).value !== 'Bracelet' ||
          worksheet.getCell(`G1`).value !== 'Movement' ||
          worksheet.getCell(`H1`).value !== 'Diameter' ||
          worksheet.getCell(`I1`).value !== 'Case Material' ||
          worksheet.getCell(`J1`).value !== 'Purchase Price' ||
          worksheet.getCell(`K1`).value !== 'Retail_EUR_RS'
        ) throw new Error(error.INVALID_COLUMN_NAME);
        break;

      case 'SwissKubik':
        // Check columns order
        if (
          worksheet.getCell(`A1`).value !== 'RMC' ||
          worksheet.getCell(`B1`).value !== 'Collection' ||
          worksheet.getCell(`C1`).value !== 'Sale Reference' ||
          worksheet.getCell(`D1`).value !== 'Material Description' ||
          worksheet.getCell(`E1`).value !== 'Description' ||
          worksheet.getCell(`F1`).value !== 'Size' ||
          worksheet.getCell(`G1`).value !== 'Materials' ||
          worksheet.getCell(`H1`).value !== 'Color' ||
          worksheet.getCell(`I1`).value !== 'Retail_EUR_RS' ||
          worksheet.getCell(`J1`).value !== 'Retail_EUR_HU'
        ) throw new Error(error.INVALID_COLUMN_NAME);
        break;

      case 'Rubber B':
        // Check columns order
        if (
          worksheet.getCell(`A1`).value !== 'RMC' ||
          worksheet.getCell(`B1`).value !== 'Sale Reference' ||
          worksheet.getCell(`C1`).value !== 'Material Description' ||
          worksheet.getCell(`D1`).value !== 'Color' ||
          worksheet.getCell(`E1`).value !== 'For Model' ||
          worksheet.getCell(`F1`).value !== 'For Clasp' ||
          worksheet.getCell(`G1`).value !== 'Purchase Price' ||
          worksheet.getCell(`H1`).value !== 'Retail_EUR_RS' ||
          worksheet.getCell(`I1`).value !== 'Retail_EUR_HU'
        ) throw new Error(error.INVALID_COLUMN_NAME);
        break;
    }

    // Set brand properties to default values
      // Rolex, Tudor
    let rmc = '';
    let collection = '';
    let productLine = '';
    let saleReference = '';
    let materialDescription = '';
    let dial = '';
    let bracelet = '';
    let box = '';
    let diameter = '';
    let caseMaterial = '';
    let exGeneveCHF = null;
    let retailRsEUR = null;
    let retailHuEUR = null;
    let retailMneEUR = null;
    let photos = [];
      // Rolex
    let braceletType = '';
    let materials = [];
      // Tudor
    let waterproofness = '';
      // Panerai
    let movement = '';
    let purchasePrice = null;
      // SwissKubik
    let description = '';
    let size = '';
    let color = '';
      // Rubber B
    let forModel = '';
    let forClasp = '';
      // Rolex - for charts
    let materialType = '';
    let chartCollection = '';

    switch (brand) {
      case 'Rolex':
        rmc = worksheet.getCell(`A${i + 2}`).value;
        collection = worksheet.getCell(`B${i + 2}`).value.toUpperCase();
        productLine = worksheet.getCell(`C${i + 2}`).value ? worksheet.getCell(`C${i + 2}`).value : '';
        saleReference = worksheet.getCell(`D${i + 2}`).value;
        materialDescription = worksheet.getCell(`E${i + 2}`).value && worksheet.getCell(`E${i + 2}`).value !== '' ? worksheet.getCell(`E${i + 2}`). value :
          worksheet.getCell(`F${i + 2}`).value && worksheet.getCell(`G${i + 2}`).value ? `${worksheet.getCell(`F${i + 2}`).value}-${worksheet.getCell(`G${i + 2}`).value}` : '';
        dial = worksheet.getCell(`F${i + 2}`).value ? worksheet.getCell(`F${i + 2}`).value :
          materialDescription && materialDescription !== '' ? materialDescription.split('-').slice(0, -1).join('-').trim() : '';
        bracelet = worksheet.getCell(`G${i + 2}`).value ? worksheet.getCell(`G${i + 2}`).value.toString() :
          materialDescription && materialDescription !== '' ? materialDescription.split('-').pop().trim(): '';
        box = worksheet.getCell(`H${i + 2}`).value ? worksheet.getCell(`H${i + 2}`).value : '';
        diameter = worksheet.getCell(`I${i + 2}`).value ? worksheet.getCell(`I${i + 2}`).value : '';
        caseMaterial = worksheet.getCell(`J${i + 2}`).value ? worksheet.getCell(`J${i + 2}`).value : dial ? setCaseMaterial(dial) : '';
        exGeneveCHF = worksheet.getCell(`K${i + 2}`).value;
        retailRsEUR = worksheet.getCell(`L${i + 2}`).value;
        retailHuEUR = worksheet.getCell(`M${i + 2}`).value;
        retailMneEUR = worksheet.getCell(`N${i + 2}`).value;
        // Additional properties
        braceletType = bracelet && bracelet !== '' ? setBraceletType(bracelet.toString()) : '';
        materials = setRolexMaterialsAndCollections(saleReference, collection).materials;
        // For charts
        materialType = setRolexMaterialsAndCollections(saleReference, collection).materialType;
        chartCollection = setRolexMaterialsAndCollections(saleReference, collection).chartCollection;
        break;

      case 'Tudor':
        rmc = worksheet.getCell(`A${i + 2}`).value;
        collection = worksheet.getCell(`B${i + 2}`).value.toUpperCase();
        productLine = worksheet.getCell(`C${i + 2}`).value ? worksheet.getCell(`C${i + 2}`).value : '';
        saleReference = worksheet.getCell(`D${i + 2}`).value;
        materialDescription = worksheet.getCell(`E${i + 2}`).value && worksheet.getCell(`E${i + 2}`).value !== '' ? worksheet.getCell(`E${i + 2}`). value :
          worksheet.getCell(`F${i + 2}`).value && worksheet.getCell(`G${i + 2}`).value ? `${worksheet.getCell(`F${i + 2}`).value}-${worksheet.getCell(`G${i + 2}`).value}` : '';
        dial = worksheet.getCell(`F${i + 2}`).value ? worksheet.getCell(`F${i + 2}`).value :
          worksheet.getCell(`E${i + 2}`).value && worksheet.getCell(`E${i + 2}`).value !== '' ? worksheet.getCell(`E${i + 2}`).value.split('-').slice(0, -1).join('-').trim() : '';
        bracelet = worksheet.getCell(`G${i + 2}`).value ? worksheet.getCell(`G${i + 2}`).value.toString() :
          worksheet.getCell(`E${i + 2}`).value && worksheet.getCell(`E${i + 2}`).value !== '' ? worksheet.getCell(`E${i + 2}`).value.split('-').pop().trim(): '';
        box = worksheet.getCell(`H${i + 2}`).value ? worksheet.getCell(`H${i + 2}`).value : '';
        diameter = worksheet.getCell(`I${i + 2}`).value ? worksheet.getCell(`I${i + 2}`).value : '';
        caseMaterial = worksheet.getCell(`J${i + 2}`).value ? worksheet.getCell(`J${i + 2}`).value : '';
        waterproofness = worksheet.getCell(`K${i + 2}`).value ? worksheet.getCell(`K${i + 2}`).value : '';
        exGeneveCHF = worksheet.getCell(`L${i + 2}`).value;
        retailRsEUR = worksheet.getCell(`M${i + 2}`).value;
        break;

      case 'Panerai':
        rmc = worksheet.getCell(`A${i + 2}`).value;
        collection = worksheet.getCell(`B${i + 2}`).value.toUpperCase();
        saleReference = worksheet.getCell(`C${i + 2}`).value;
        materialDescription = worksheet.getCell(`D${i + 2}`).value ? worksheet.getCell(`D${i + 2}`).value : '';
        dial = worksheet.getCell(`E${i + 2}`).value ? worksheet.getCell(`E${i + 2}`).value : '';
        bracelet = worksheet.getCell(`F${i + 2}`).value ? worksheet.getCell(`F${i + 2}`).value : '';
        box = worksheet.getCell(`G${i + 2}`).value ? worksheet.getCell(`G${i + 2}`).value : '';
        diameter = worksheet.getCell(`H${i + 2}`).value ? worksheet.getCell(`H${i + 2}`).value : '';
        caseMaterial = worksheet.getCell(`I${i + 2}`).value ? worksheet.getCell(`I${i + 2}`).value : '';
        purchasePrice = worksheet.getCell(`J${i + 2}`).value;
        retailRsEUR = worksheet.getCell(`K${i + 2}`).value;
        break;

      case 'SwissKubik':
        rmc = worksheet.getCell(`A${i + 2}`).value;
        collection = worksheet.getCell(`B${i + 2}`).value.toUpperCase();
        saleReference = worksheet.getCell(`C${i + 2}`).value;
        materialDescription = worksheet.getCell(`D${i + 2}`).value ? worksheet.getCell(`D${i + 2}`).value : '';
        description = worksheet.getCell(`E${i + 2}`).value ? worksheet.getCell(`E${i + 2}`).value : '';
        size = worksheet.getCell(`F${i + 2}`).value ? worksheet.getCell(`F${i + 2}`).value : '';
        materials = worksheet.getCell(`G${i + 2}`).value ? worksheet.getCell(`G${i + 2}`).value.split(',').map(el => el.trim()).map(el => el ? el[0].toUpperCase() + el.substring(1) : '').filter(el => el !== '') : [];
        color = worksheet.getCell(`H${i + 2}`).value ? worksheet.getCell(`H${i + 2}`).value : '';
        retailRsEUR = worksheet.getCell(`I${i + 2}`).value;
        retailHuEUR = worksheet.getCell(`J${i + 2}`).value;

        // Validate SwissKubik materials
        for (let el of materials) if (!jewelryMaterials.includes(el)) throw new Error(error.INVALID_MATERIAL);
        break;

      case 'Rubber B':
        rmc = worksheet.getCell(`A${i + 2}`).value;
        saleReference = worksheet.getCell(`B${i + 2}`).value;
        materialDescription = worksheet.getCell(`C${i + 2}`).value ? worksheet.getCell(`C${i + 2}`).value : '';
        color = worksheet.getCell(`D${i + 2}`).value ? worksheet.getCell(`D${i + 2}`).value : '';
        forModel = worksheet.getCell(`E${i + 2}`).value ? worksheet.getCell(`E${i + 2}`).value : '';
        forClasp = worksheet.getCell(`F${i + 2}`).value ? worksheet.getCell(`F${i + 2}`).value : '';
        purchasePrice = worksheet.getCell(`G${i + 2}`).value;
        retailRsEUR = worksheet.getCell(`H${i + 2}`).value;
        retailHuEUR = worksheet.getCell(`I${i + 2}`).value;
        // Additional properties
        collection = worksheet.getCell(`B${i + 2}`).value.toUpperCase();
        break;
    }

    // Photos
    photos = [`${rmc}.${extension}`];

    // Validate required properties
    if (!rmc || !collection || !saleReference) throw new Error(error.MISSING_PARAMETERS);
    if (!retailRsEUR) throw new Error(error.MISSING_PRICE);
    if (['Rolex', 'SwissKubik', 'Rubber B'].includes(brand) && !retailHuEUR) throw new Error(error.MISSING_PRICE);
    if (['Panerai', 'Rubber B'].includes(brand) && !purchasePrice) throw new Error(error.MISSING_PURCHASE_PRICE);

    // Push product rmc into 'rmcArray'
    rmcArray.push(rmc);

    // Create 'boutiques' array
    let boutiques = [];

    // Check if uploaded product already exists in DB
    const existingProduct = existingProducts.find(obj => obj.basicInfo.rmc === rmc);

    // Scenario 1: If uploaded product doesn't exist in DB, Create new modifiedProduct for the purpose of reviewing and also (if 'save' is sent as 'true') Create new product to be stored in DB
    if (!existingProduct) {
      // For each existing store create object containing: store ID, store name, respective price in EUR, local price, priceHistory and VAT, and push the object into 'boutiques' array
      for (const store of existingStores) {
        if (store.name === 'Belgrade') boutiques.push({
          store: store._id, storeName: store.name,
          price: retailRsEUR,
          VATpercent: vatRs,
          priceLocal: Math.ceil(retailRsEUR * rateRSDWatches / 1000) * 1000,
          priceHistory: [{
            date: new Date(),
            price: retailRsEUR,
            VAT: vatRs,
            priceLocal: Math.ceil(retailRsEUR * rateRSDWatches / 1000) * 1000
          }]
        });
        if (store.name === 'Budapest' && ['Rolex', 'SwissKubik', 'Rubber B'].includes(brand)) boutiques.push({
          store: store._id, storeName: store.name,
          price: retailHuEUR,
          VATpercent: vatHu,
          priceLocal: Math.ceil(retailHuEUR * rateHUFWatches / 1000) * 1000,
          priceHistory: [{
            date: new Date(),
            price: retailHuEUR,
            VAT: vatHu,
            priceLocal: Math.ceil(retailHuEUR * rateHUFWatches / 1000) * 1000
          }]
        });
        if (store.name === 'Porto Montenegro' && brand === 'Rolex') boutiques.push({
          store: store._id,
          storeName: store.name,
          price: retailMneEUR,
          VATpercent: vatMn,
          priceHistory: [{
            date: new Date(),
            price: retailMneEUR,
            VAT: vatMn
          }]
        });
      }

      // 3. Create modified product
      const newModifiedProduct = new ModifiedProduct({
        status: 'new',
        brand,
        boutiques,
        basicInfo: {
          rmc,
          collection,
          productLine,
          saleReference,
          materialDescription,
          dial,
          bracelet,
          box,
          diameter,
          caseMaterial,
          exGeneveCHF,
          braceletType,
          materials,
          photos,
          // Tudor
          waterproofness,
          // Panerai
          movement,
          purchasePrice,
          // SwissKubik
          description,
          size,
          color,
          // Rubber B
          forModel,
          forClasp
        }
      });

      // Push created modified product to 'modifiedProducts' array
      modifiedProducts.push(newModifiedProduct);
      // z++;
      // console.log(`${z}. new modified product: ${rmc}`);

      // If 'save' is true
      if (save) {
        // 4. Create new product
        const newProduct = new Product({
          status: 'new',
          brand,
          boutiques,
          basicInfo: {
            rmc,
            collection,
            productLine,
            saleReference,
            materialDescription,
            dial,
            bracelet,
            box,
            diameter,
            caseMaterial,
            exGeneveCHF,
            braceletType,
            materials,
            photos,
            // Tudor
            waterproofness,
            // Panerai
            movement,
            purchasePrice,
            // SwissKubik
            description,
            size,
            color,
            // Rubber B
            forModel,
            forClasp,
            // Rolex
            materialType,
            chartCollection
          }
        });

        // Push created product to 'newProducts' array
        newProducts.push(newProduct);
        // j++;
        // console.log(`${j}. new product: ${rmc} - saved to DB`);
      }
    }

    // Scenario 2: If the uploaded product (with the same RMC) already exists in DB -> Update it with new information
    if (existingProduct) {
      // Copy content of 'existingProduct' 'boutiques' array of objects
      let boutiques2 = JSON.parse(JSON.stringify(existingProduct.boutiques));

      // 2.A Set new prices and push these prices to 'priceHistory' array -> within each belonging boutique object of 'boutiques2' array
      for (let boutique of boutiques2) {
        if (boutique.storeName === 'Belgrade') {
          if (retailRsEUR) {
            boutique.price = retailRsEUR,
            boutique.VATpercent = vatRs;
            boutique.priceLocal = Math.ceil(retailRsEUR * rateRSDWatches / 1000) * 1000;
            boutique.priceHistory.push({
              date: new Date(),
              price: retailRsEUR,
              VAT: vatRs,
              priceLocal: Math.ceil(retailRsEUR * rateRSDWatches / 1000) * 1000
            });
          }
        }
        if (boutique.storeName === 'Budapest' && ['Rolex', 'SwissKubik', 'Rubber B'].includes(brand)) {
          if (retailHuEUR) {
            boutique.price = retailHuEUR;
            boutique.VATpercent = vatHu;
            boutique.priceLocal = Math.ceil(retailHuEUR * rateHUFWatches / 1000) * 1000;
            boutique.priceHistory.push({
              date: new Date(),
              price: retailHuEUR,
              VAT: vatHu,
              priceLocal: Math.ceil(retailHuEUR * rateHUFWatches / 1000) * 1000
            });
          }
        }
        if (boutique.storeName === 'Porto Montenegro' && brand === 'Rolex') {
          if (retailMneEUR) {
            boutique.price = retailMneEUR;
            boutique.VATpercent = vatMn;
            boutique.priceHistory.push({
              date: new Date(),
              price: retailMneEUR,
              VAT: vatMn
            });
          }
        }
      }

      // 2.B Create 'updateSet' object
      let updateSet = { boutiques: boutiques2 };

      // All brands (3)
      if (existingProduct.basicInfo.collection != collection) updateSet['basicInfo.collection'] = collection;
      if (existingProduct.basicInfo.saleReference != saleReference) updateSet['basicInfo.saleReference'] = saleReference;
      if (existingProduct.basicInfo.materialDescription != materialDescription) updateSet['basicInfo.materialDescription'] = materialDescription;

      // Watches properties (4)
      if (['Rolex', 'Tudor', 'Panerai'].includes(brand)) {
        if (existingProduct.basicInfo.dial != dial) updateSet['basicInfo.dial'] = dial;
        if (existingProduct.basicInfo.bracelet != bracelet) updateSet['basicInfo.bracelet'] = bracelet;
        if (existingProduct.basicInfo.diameter != diameter) updateSet['basicInfo.diameter'] = diameter;
        if (existingProduct.basicInfo.caseMaterial != caseMaterial) updateSet['basicInfo.caseMaterial'] = caseMaterial;
      }

      // Rolex and Tudor properties (3)
      if (['Rolex', 'Tudor'].includes(brand)) {
        if (existingProduct.basicInfo.productLine != productLine) updateSet['basicInfo.productLine'] = productLine;
        if (existingProduct.basicInfo.box != box) updateSet['basicInfo.box'] = box;
        if (existingProduct.basicInfo.exGeneveCHF != exGeneveCHF) updateSet['basicInfo.exGeneveCHF'] = exGeneveCHF;
      }

      // Rolex properties (3)
      if (['Rolex'].includes(brand)) {
        if (existingProduct.basicInfo.braceletType != braceletType) updateSet['basicInfo.braceletType'] = braceletType;
        if (existingProduct.basicInfo.materialType != materialType) updateSet['basicInfo.materialType'] = materialType;
        if (existingProduct.basicInfo.chartCollection != chartCollection) updateSet['basicInfo.chartCollection'] = chartCollection;
      }

      // Rolex and SwissKubik properties (1)
      if (['Rolex', 'SwissKubik'].includes(brand)) {
        if (_.isEqual(existingProduct.basicInfo.materials, materials) === false) updateSet['basicInfo.materials'] = materials;
      }

      // Tudor properties (1)
      if (['Tudor'].includes(brand)) {
        if (existingProduct.basicInfo.waterproofness != waterproofness) updateSet['basicInfo.waterproofness'] = waterproofness;
      }

      // Panerai and Rubber B properties (1)
      if (['Rolex', 'Rubber B'].includes(brand)) {
        if (existingProduct.basicInfo.purchasePrice != purchasePrice) updateSet['basicInfo.purchasePrice'] = purchasePrice;
      }

      // Panerai properties (1)
      if (['Panerai'].includes(brand)) {
        if (existingProduct.basicInfo.movement != movement) updateSet['basicInfo.movement'] = movement;
      }

      // SwissKubik and Rubber B properties (1)
      if (['SwissKubik', 'Rubber B'].includes(brand)) {
        if (existingProduct.basicInfo.color != color) updateSet['basicInfo.color'] = color;
      }

      // SwissKubik properties (2)
      if (['SwissKubik'].includes(brand)) {
        if (existingProduct.basicInfo.size != size) updateSet['basicInfo.size'] = size;
        if (existingProduct.basicInfo.description != description) updateSet['basicInfo.description'] = description;
      }

      // Rubber B properties (2)
      if (['Rubber B'].includes(brand)) {
        if (existingProduct.basicInfo.forModel != forModel) updateSet['basicInfo.forModel'] = forModel;
        if (existingProduct.basicInfo.forClasp != forClasp) updateSet['basicInfo.forClasp'] = forClasp;
      }

      // Scenario 2.1: If any of retail prices of the existing product has been changed or any other property has been changed -> then change 'status' of the product model to 'changed' and create modified product
      if (existingProduct.status !== 'deleted' &&
        // Can't find boutique in 'boutiques' array whose 'retailRsEur' price is equal to the sent one of the uploaded product
        existingProduct.boutiques.find(boutique => boutique.price === retailRsEUR) === undefined ||
        // Can't find boutique in 'boutiques' array whose 'retailHuEUR' price is equal to the sent one of the uploaded product
        (['Rolex', 'SwissKubik', 'Rubber B'].includes(brand) && existingProduct.boutiques.find(boutique => boutique.price === retailHuEUR) === undefined) ||
        // Can't find boutique in 'boutiques' array whose 'retailMneEUR' price is equal to the sent one of the uploaded product
        (brand === 'Rolex' && existingProduct.boutiques.find(boutique => boutique.price === retailMneEUR) === undefined) ||

        // 'updateSet' is not empty
        Object.keys(updateSet).length !== 0
      ) {
        updateSet.status = 'changed';

        // 3. Create modified product -> to be used for review
        const newModifiedProduct = new ModifiedProduct({
          status: 'changed',
          brand,
          boutiques: boutiques2,
          basicInfo: {
            rmc,
            collection,
            productLine,
            saleReference,
            materialDescription,
            dial,
            bracelet,
            box,
            diameter,
            caseMaterial,
            exGeneveCHF,
            braceletType,
            materials,
            photos,
            // Tudor
            waterproofness,
            // Panerai
            movement,
            purchasePrice,
            // SwissKubik
            description,
            size,
            color,
            // Rubber B
            forModel,
            forClasp
          }
        });

        // Push created modified product to 'modifiedProducts' array
        modifiedProducts.push(newModifiedProduct);
        // k++;
        // console.log(`${k}. updated modified product: ${rmc}`);
      }

      // Scenario 2.2: If existing product's 'status' is equal to 'deleted' -> then change 'status' of the product model to 'restored'
      if (existingProduct.status === 'deleted') {
        updateSet.status = 'restored';

        // 3. Create a new modified product for review
        const newModifiedProduct = new ModifiedProduct({
          status: 'restored',
          brand,
          boutiques: boutiques2,
          basicInfo: {
            rmc,
            collection,
            productLine,
            saleReference,
            materialDescription,
            dial,
            bracelet,
            box,
            diameter,
            caseMaterial,
            exGeneveCHF,
            braceletType,
            materials,
            photos,
            // Tudor
            waterproofness,
            // Panerai
            movement,
            purchasePrice,
            // SwissKubik
            description,
            size,
            color,
            // Rubber B
            forModel,
            forClasp
          }
        });

        // Push created modified product to 'modifiedProducts' array
        modifiedProducts.push(newModifiedProduct);
        // l++;
        // console.log(`${l}. restored modified product: ${rmc}`);
      }

      // If 'save' is true
      if (save) {
        // Push updated product to 'updatedProducts' array
        updatedProducts.push(
          // 4. Update product -> change it's prices (only prices and 'status' are changed)
          Product.updateOne(
            { 'basicInfo.rmc': rmc },
            { $set: updateSet },
            { new: true }
          )
        );
        // m++;
        // console.log(`${m}. updated product: ${rmc} - changes saved to DB`);
      }
    }
  }

  // // Scenario 3: Check if there are products in DB that are not included in the newly uploaded excel RMC table -> get array of IDs of such products
  // const toDelete = await Product.distinct('_id', { 'basicInfo.rmc': { $nin: rmcArray }, status: { $ne: 'deleted' } });

  // // For all products that exist in DB but are not included in the uploaded excel RMC table
  // for (const productId of toDelete) {
  //   // Find product in DB
  //   const deletedProduct = await Product.findOne({ _id: productId }).lean();

  //   // Create new modified product (out of the not included product) for review
  //   const newModifiedProduct = new ModifiedProduct({
  //     status: 'deleted',
  //     brand,
  //     boutiques: deletedProduct.boutiques,
  //     basicInfo: {
  //       rmc: deletedProduct.basicInfo.rmc,
  //       collection: deletedProduct.basicInfo.collection,
  //       productLine: deletedProduct.basicInfo.productLine,
  //       saleReference: deletedProduct.basicInfo.saleReference,
  //       materialDescription: deletedProduct.basicInfo.materialDescription,
  //       dial: deletedProduct.basicInfo.dial,
  //       bracelet: deletedProduct.basicInfo.bracelet,
  //       box: deletedProduct.basicInfo.box,
  //       exGeneveCHF: deletedProduct.basicInfo.exGeneveCHF,
  //       diameter: deletedProduct.basicInfo.diameter,
  //       photos: deletedProduct.basicInfo.photos,
  //     }
  //   });

  //   // Push created modified product object to 'modifiedProducts' array
  //   modifiedProducts.push(newModifiedProduct);

  //   // If 'save' is true
  //   if (save) {
  //     // Push not included product to 'updatedProducts' array
  //     updatedProducts.push(
  //       // Update product -> change 'status' of not included product to 'deleted', but do not actually delete it from DB
  //       Product.updateOne(
  //         { _id: deletedProduct._id },
  //         { status: 'deleted' },
  //         { new: true }
  //       )
  //     );
  //   }
  // }

  // Insert all modified products into DB
  await ModifiedProduct.insertMany(modifiedProducts);

  // If 'save' is true -> execute creation of new products and update of products
  if (save) {
    await Promise.all([
      // Update all 'updatedProducts'
      Promise.all(updatedProducts),
      // Save all 'newProducts'
      Product.insertMany(newProducts)
    ]);
  }

  return res.status(201).send({
    message: 'Successfully imported new product models',
  });
};

/**
 * @api {post} /product/soonInStock Import soonInStocks
 * @apiVersion 1.0.0
 * @apiName importSoonInStocks
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (formData) {Binary} file Binary file uploaded
 * @apiParam (query) {String} boutique Store ID
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully imported soon in stock watches",
 }
 *
 * @apiUse MissingParamsError
 * @apiUse MissingStatus
 * @apiUse MissingLocation
 * @apiUse MissingExGenevaPrice
 * @apiUse MissingShipmentDate
 * @apiUse InvalidStatus
 * @apiUse InvalidValue
 * @apiUse InvalidColumnName
 * @apiUse InvalidShipmentDate
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
 module.exports.importSoonInStocks = async (req, res) => {
  const { _id: userId } = req.user;
  const { objectName } = req.file;
  const { boutique } = req.query;

  // Check if req.file has been sent
  if (!objectName || !boutique) throw new Error(error.MISSING_PARAMETERS);

  // Create a stream to pull excel file from the Minion server
  const dataStream = await minioClient.getObject(environments.MINIO_BUCKET, objectName);

  // Create an empty workbook
  const workbook = new exceljs.Workbook();

  // Fill in the workbook with the excel data pulled from the Minion server
  await workbook.xlsx.read(dataStream);

  // Set 'worksheet' constant -> that contains data from the 1st sheet of 'workbook'
  const worksheet = workbook.getWorksheet(1);

  // Get list of existing stores in DB
  const existingStores = await Store.find().lean();

  // Get list of existing product models in DB
  const existingProducts = await Product.find({ brand: { $in: ['Rolex', 'Tudor'] } }).lean();

  // Check if store
  const store = existingStores.find(obj => obj._id.toString() === boutique);
  if (!store) throw new Error(error.INVALID_VALUE);

  // Create 'newWatches' array
  const newWatches = [];
  const newActivities = [];
  const updateActivities = [];

  // Check columns order
  if (
    worksheet.getCell(`A1`).value !== 'PGP' ||
    worksheet.getCell(`B1`).value !== 'Location' ||
    worksheet.getCell(`C1`).value !== 'Status' ||
    worksheet.getCell(`D1`).value !== 'Reserved For - Client ID' ||
    worksheet.getCell(`E1`).value !== 'Reservation Time' ||
    worksheet.getCell(`F1`).value !== 'Comment' ||
    worksheet.getCell(`G1`).value !== 'Previous Serial' ||
    worksheet.getCell(`H1`).value !== 'soldToParty' ||
    worksheet.getCell(`I1`).value !== 'shipToParty' ||
    worksheet.getCell(`J1`).value !== 'billToParty' ||
    worksheet.getCell(`K1`).value !== 'packingNumber' ||
    worksheet.getCell(`L1`).value !== 'invoiceDate' ||
    worksheet.getCell(`M1`).value !== 'invoiceNumber' ||
    worksheet.getCell(`N1`).value !== 'Serial' ||
    worksheet.getCell(`O1`).value !== 'boxCode' ||
    worksheet.getCell(`P1`).value !== 'exGenevaPrice' ||
    worksheet.getCell(`Q1`).value !== 'sectorId' ||
    worksheet.getCell(`R1`).value !== 'sector' ||
    worksheet.getCell(`S1`).value !== 'boitId' ||
    worksheet.getCell(`T1`).value !== 'boitRef' ||
    worksheet.getCell(`U1`).value !== 'caspId' ||
    worksheet.getCell(`V1`).value !== 'cadranDesc' ||
    worksheet.getCell(`W1`).value !== 'ldisId' ||
    worksheet.getCell(`X1`).value !== 'ldisq' ||
    worksheet.getCell(`Y1`).value !== 'brspId' ||
    worksheet.getCell(`Z1`).value !== 'brspRef' ||
    worksheet.getCell(`AA1`).value !== 'prixSuisse' ||
    worksheet.getCell(`AB1`).value !== 'RMC' ||
    worksheet.getCell(`AC1`).value !== 'umIntern' ||
    worksheet.getCell(`AD1`).value !== 'umExtern' ||
    worksheet.getCell(`AE1`).value !== 'rfid' ||
    worksheet.getCell(`AF1`).value !== 'packingType' ||
    worksheet.getCell(`AG1`).value !== 'Shipment Date (YYYY-MM-DD)'
  ) throw new Error(error.INVALID_COLUMN_NAME);

  let j = 0;

  // For each row in the 'Tppack' table
  for (let i = 0; i < worksheet.actualRowCount - 1; i++) {
    // Get editable watch features
    const pgpReference = worksheet.getCell(`A${i + 2}`).value;
    const location = worksheet.getCell(`B${i + 2}`).value;
    const status = worksheet.getCell(`C${i + 2}`).value;
    const reservedFor = worksheet.getCell(`D${i + 2}`).value;
    const reservationTime = worksheet.getCell(`E${i + 2}`).value;
    const comment = worksheet.getCell(`F${i + 2}`).value ? worksheet.getCell(`F${i + 2}`).value : '';
    const previousSerial = worksheet.getCell(`G${i + 2}`).value;

    // Get 'rmc' and 'serialNumber'
    const rmc = worksheet.getCell(`AB${i + 2}`).value;
    const serialNumber = worksheet.getCell(`N${i + 2}`).value;

    // Get other watch features:
    const soldToParty = worksheet.getCell(`H${i + 2}`).value ? worksheet.getCell(`H${i + 2}`).value : '';
    const shipToParty = worksheet.getCell(`I${i + 2}`).value ? worksheet.getCell(`I${i + 2}`).value : '';
    const card = worksheet.getCell(`I${i + 2}`).value ? worksheet.getCell(`I${i + 2}`).value : '';
    const billToParty = worksheet.getCell(`J${i + 2}`).value ? worksheet.getCell(`J${i + 2}`).value : '';
    const packingNumber = worksheet.getCell(`K${i + 2}`).value ? worksheet.getCell(`K${i + 2}`).value : '';
    const dateOfInvoice = worksheet.getCell(`L${i + 2}`).value.toString();
    const invoiceDate = dateOfInvoice.substring(0, 4) + "-" + dateOfInvoice.substring(4, 6) + "-" + dateOfInvoice.substring(6);
    const invoiceNumber = worksheet.getCell(`M${i + 2}`).value ? worksheet.getCell(`M${i + 2}`).value : '';
    const boxCode = worksheet.getCell(`O${i + 2}`).value ? worksheet.getCell(`O${i + 2}`).value : '';
    const exGenevaPrice = worksheet.getCell(`P${i + 2}`).value;
    const sectorId = worksheet.getCell(`Q${i + 2}`).value ? worksheet.getCell(`Q${i + 2}`).value : '';
    const sector = worksheet.getCell(`R${i + 2}`).value ? worksheet.getCell(`R${i + 2}`).value : '';
    const boitId = worksheet.getCell(`S${i + 2}`).value ? worksheet.getCell(`S${i + 2}`).value : '';
    const boitRef = worksheet.getCell(`T${i + 2}`).value ? worksheet.getCell(`T${i + 2}`).value : '';
    const caspId = worksheet.getCell(`U${i + 2}`).value ? worksheet.getCell(`U${i + 2}`).value : '';
    const cadranDesc = worksheet.getCell(`V${i + 2}`).value ? worksheet.getCell(`V${i + 2}`).value : '';
    const ldisId = worksheet.getCell(`W${i + 2}`).value;
    const ldisq = worksheet.getCell(`X${i + 2}`).value ? worksheet.getCell(`X${i + 2}`).value : '';
    const brspId = worksheet.getCell(`Y${i + 2}`).value ? worksheet.getCell(`Y${i + 2}`).value : '';
    const brspRef = worksheet.getCell(`Z${i + 2}`).value ? worksheet.getCell(`Z${i + 2}`).value : '';
    const prixSuisse = worksheet.getCell(`AA${i + 2}`).value;
    const umIntern = worksheet.getCell(`AC${i + 2}`).value ? worksheet.getCell(`AC${i + 2}`).value : '';
    const umExtern = worksheet.getCell(`AD${i + 2}`).value ? worksheet.getCell(`AD${i + 2}`).value : '';
    const rfid = worksheet.getCell(`AE${i + 2}`).value ? worksheet.getCell(`AE${i + 2}`).value : '';
    const packingType = worksheet.getCell(`AF${i + 2}`).value ? worksheet.getCell(`AF${i + 2}`).value : '';

    // Shipment date
    const shipmentDate = worksheet.getCell(`AG${i + 2}`).value ? worksheet.getCell(`AG${i + 2}`).value : '';

    // Check required properties
    if (!rmc || !serialNumber || !pgpReference) throw new Error(error.MISSING_PARAMETERS);
    if (!status) throw new Error(error.MISSING_STATUS);
    if (!location) throw new Error(error.MISSING_LOCATION);
    if (!exGenevaPrice) throw new Error(error.MISSING_EX_GENEVA_PRICE);
    if (!shipmentDate) throw new Error(error.MISSING_SHIPMENT_DATE);

    // Validate status
    if (status && !statuses.includes(status)) throw new Error(error.INVALID_STATUS);

    // Validate 'shipmentDate'
    if (shipmentDate.length !== 10 || shipmentDate.split('-')[0].length !== 4 || shipmentDate.split('-')[1].length !== 2 || shipmentDate.split('-')[2].length !== 2 || !moment(shipmentDate, 'YYYY-MM-DD').isValid()) throw new Error(error.INVALID_SHIPMENT_DATE);

    // Find product in DB with the sent 'rmc'
    const product = existingProducts.find(obj => obj.basicInfo.rmc === rmc);

    // Check if product was found
    if (!product) throw new Error(error.NOT_FOUND);

    // 1. Create new soonInStock watch
    const soonInStock = new SoonInStock({
      // Product RMC model features
      rmc,
      product: product._id,
      brand: product.brand,
      dial: product.basicInfo.dial,
      // Watch features
      serialNumber,
      pgpReference,
      // Changeable properties
      store: boutique,
      status,
      location,
      reservedFor,
      reservationTime,
      comment,
      origin: 'Geneva',
      card,
      // Other features
      soldToParty,
      shipToParty,
      billToParty,
      packingNumber,
      invoiceDate,
      invoiceNumber,
      boxCode,
      exGenevaPrice,
      sectorId,
      sector,
      boitId,
      boitRef,
      caspId,
      cadranDesc,
      ldisId,
      ldisq,
      brspId,
      brspRef,
      prixSuisse,
      umIntern,
      umExtern,
      rfid,
      packingType,
      shipmentDate
    });

    // Push created soonInStock watch to 'newWatches' array
    newWatches.push(soonInStock);

    j++;
    // console.log(`${j}. imported soon in stock watch: ${rmc} - ${serialNumber}`);

    // 2. Create new activity
    const logComment = `Entered as soon in stock with serial number: '${serialNumber}'`;
    const newActivity = createActivity('Product', userId, null, logComment, null, soonInStock._id, serialNumber, null, new Date());

    // Push created activity to 'newActivities' array
    newActivities.push(newActivity.save());

    // console.log(`created activity with serial number: '${serialNumber}'`);

    // 3. Update previously created activities
    const updateActivity = Activity.updateMany(
      { serialNumber: previousSerial },
      { $set: { serialNumber } }
    );

    // Push to array
    updateActivities.push(updateActivity);
  }

  // Execute
  await Promise.all([
    SoonInStock.insertMany(newWatches),
    ...newActivities,
    ...updateActivities
  ]);

  return res.status(200).send({
    message: 'Successfully imported soon in stock watches',
  });
};

/**
 * @api {post} /product/stock/excel Import stock
 * @apiVersion 1.0.0
 * @apiName importStock
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (formData) {Binary} file Binary file uploaded
 * @apiParam (query) {String} storeId Store ID
 * @apiParam (query) {String='Panerai', 'SwissKubik', 'Rubber B', 'Messika', 'Roberto Coin', 'Petrovic Diamonds'} brand Brand
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 201 OK
 {
   "message": "Successfully imported new stock",
 }
 *
 * @apiUse MissingParamsError
 * @apiUse InvalidValue
 * @apiUse InvalidColumnName
 * @apiUse MissingPrice
 * @apiUse MissingExGenevaPrice
 * @apiUse MissingInvoiceDate
 * @apiUse InvalidStockDate
 * @apiUse InvalidInvoiceDate
 * @apiUse MissingJewelryType
 * @apiUse MissingLocation
 * @apiUse MissingPurchasePrice
 * @apiUse InvalidSize
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
 module.exports.importStock = async (req, res) => {
  const { _id: userId } = req.user;
  const { objectName } = req.file;
  const { storeId, brand } = req.query;

  // Check if req.file has been sent
  if (!objectName || !storeId || !brand) throw new Error(error.MISSING_PARAMETERS);

  // Validate brand
  if (!['Panerai', 'SwissKubik', 'Rubber B', 'Messika', 'Roberto Coin', 'Petrovic Diamonds'].includes(brand)) throw new Error(error.INVALID_VALUE);

  // Create a stream to pull excel file from the Minion server
  const dataStream = await minioClient.getObject(environments.MINIO_BUCKET, objectName);

  // Create an empty workbook
  const workbook = new exceljs.Workbook();

  // Fill in the workbook with the excel data pulled from the Minion server
  await workbook.xlsx.read(dataStream);

  // Set 'worksheet' constant -> that contains data from the 1st sheet of 'workbook'
  const worksheet = workbook.getWorksheet(1);

  // Get list of existing stores in DB
  const existingStores = await Store.find().lean();

  // Get list of existing product models in DB
  const existingProducts = await Product.find({ brand }).lean();

  // Check if store exists in DB and validate it
  const store = existingStores.find(obj => obj._id.toString() === storeId);
  if (!store) throw new Error(error.INVALID_VALUE);

  // Create 'newWatches' array
  const addToStocks = [];
  const newProducts = [];
  const newActivities = [];

  switch (brand) {
    case 'Panerai':
      if (
        worksheet.getCell(`A1`).value !== 'Sale Reference' ||
        worksheet.getCell(`B1`).value !== 'Serial Number' ||
        worksheet.getCell(`C1`).value !== 'PGP' ||
        worksheet.getCell(`D1`).value !== 'Location' ||
        worksheet.getCell(`E1`).value !== 'Comment' ||
        worksheet.getCell(`F1`).value !== 'Stock Date' ||
        worksheet.getCell(`G1`).value !== 'Invoice Date' ||
        worksheet.getCell(`H1`).value !== 'Purchase Price (RSD/HUF)'
      ) throw new Error(error.INVALID_COLUMN_NAME);
      break;



    case 'SwissKubik':
      if (
        worksheet.getCell(`A1`).value !== 'Sale Reference' ||
        worksheet.getCell(`B1`).value !== 'Serial Number' ||
        worksheet.getCell(`C1`).value !== 'PGP' ||
        worksheet.getCell(`D1`).value !== 'Location' ||
        worksheet.getCell(`E1`).value !== 'Comment' ||
        worksheet.getCell(`F1`).value !== 'Stock Date' ||
        worksheet.getCell(`G1`).value !== 'Invoice Number' ||
        worksheet.getCell(`H1`).value !== 'Invoice Date' ||
        worksheet.getCell(`I1`).value !== 'Ex Geneva Price (*100)' ||
        worksheet.getCell(`J1`).value !== 'Purchase Price (RSD/HUF)'
      ) throw new Error(error.INVALID_COLUMN_NAME);
      break;

    case 'Rubber B':
      if (
        worksheet.getCell(`A1`).value !== 'RMC' ||
        worksheet.getCell(`B1`).value !== 'Serial Number' ||
        worksheet.getCell(`C1`).value !== 'PGP' ||
        worksheet.getCell(`D1`).value !== 'Location' ||
        worksheet.getCell(`E1`).value !== 'Size' ||
        worksheet.getCell(`F1`).value !== 'Comment' ||
        worksheet.getCell(`G1`).value !== 'Stock Date' ||
        worksheet.getCell(`H1`).value !== 'Invoice Date' ||
        worksheet.getCell(`I1`).value !== 'Purchase Price (RSD/HUF)'
      ) throw new Error(error.INVALID_COLUMN_NAME);
      break;

    case 'Messika':
      if (
        worksheet.getCell(`A1`).value !== 'Collection' ||
        worksheet.getCell(`B1`).value !== 'Sale Reference' ||
        worksheet.getCell(`C1`).value !== 'Material Description' ||
        worksheet.getCell(`D1`).value !== 'Jewelry Type' ||
        worksheet.getCell(`E1`).value !== 'Size' ||
        worksheet.getCell(`F1`).value !== 'Gold Weight' ||
        worksheet.getCell(`G1`).value !== 'Stones Weight' ||
        worksheet.getCell(`H1`).value !== 'Stones Qty' ||
        worksheet.getCell(`I1`).value !== 'Brilliants' ||
        worksheet.getCell(`J1`).value !== 'Purchase Price' ||
        worksheet.getCell(`K1`).value !== 'Purchase Price (RSD/HUF)' ||
        worksheet.getCell(`L1`).value !== 'Retail Price' ||
        worksheet.getCell(`M1`).value !== 'Serial Number' ||
        worksheet.getCell(`N1`).value !== 'PGP' ||
        worksheet.getCell(`O1`).value !== 'Location' ||
        worksheet.getCell(`P1`).value !== 'Comment' ||
        worksheet.getCell(`Q1`).value !== 'Stock Date' ||
        worksheet.getCell(`R1`).value !== 'Invoice Number' ||
        worksheet.getCell(`S1`).value !== 'Invoice Date'
      ) throw new Error(error.INVALID_COLUMN_NAME);
      break;

    case 'Roberto Coin':
      if (
        worksheet.getCell(`A1`).value !== 'Collection' ||
        worksheet.getCell(`B1`).value !== 'Sale Reference' ||
        worksheet.getCell(`C1`).value !== 'Jewelry Type' ||
        worksheet.getCell(`D1`).value !== 'Size' ||
        worksheet.getCell(`E1`).value !== 'Gold Color' ||
        worksheet.getCell(`F1`).value !== 'Gold Weight' ||
        worksheet.getCell(`G1`).value !== 'Dia Gia' ||
        worksheet.getCell(`H1`).value !== 'Dia Qty' ||
        worksheet.getCell(`I1`).value !== 'Dia Carat' ||
        worksheet.getCell(`J1`).value !== 'Ruby Qty' ||
        worksheet.getCell(`K1`).value !== 'Ruby Carat' ||
        worksheet.getCell(`L1`).value !== 'Stones' ||
        worksheet.getCell(`M1`).value !== 'Purchase Price' ||
        worksheet.getCell(`N1`).value !== 'Purchase Price (RSD/HUF)' ||
        worksheet.getCell(`O1`).value !== 'Retail Price' ||
        worksheet.getCell(`P1`).value !== 'Serial Number' ||
        worksheet.getCell(`Q1`).value !== 'PGP' ||
        worksheet.getCell(`R1`).value !== 'Location' ||
        worksheet.getCell(`S1`).value !== 'Comment' ||
        worksheet.getCell(`T1`).value !== 'Stock Date' ||
        worksheet.getCell(`U1`).value !== 'Invoice Number' ||
        worksheet.getCell(`V1`).value !== 'Invoice Date'
      ) throw new Error(error.INVALID_COLUMN_NAME);
      break;

    case 'Petrovic Diamonds':
      if (
        worksheet.getCell(`A1`).value !== 'Sale Reference' ||
        worksheet.getCell(`B1`).value !== 'Jewelry Type' ||
        worksheet.getCell(`C1`).value !== 'Size' ||
        worksheet.getCell(`D1`).value !== 'Gold Weight' ||
        worksheet.getCell(`E1`).value !== 'Dia 1 Carat' ||
        worksheet.getCell(`F1`).value !== 'Dia 1 Color' ||
        worksheet.getCell(`G1`).value !== 'Dia 1 Clarity' ||
        worksheet.getCell(`H1`).value !== 'Dia 1 Shape' ||
        worksheet.getCell(`I1`).value !== 'Dia 1 Cut' ||
        worksheet.getCell(`J1`).value !== 'Dia 1 Polish' ||
        worksheet.getCell(`K1`).value !== 'Dia 1 Symmetry' ||
        worksheet.getCell(`L1`).value !== 'Dia 1 Certificate' ||
        worksheet.getCell(`M1`).value !== 'Dia 2 Carat' ||
        worksheet.getCell(`N1`).value !== 'Dia 2 Color' ||
        worksheet.getCell(`O1`).value !== 'Dia 2 Clarity' ||
        worksheet.getCell(`P1`).value !== 'Dia 2 Shape' ||
        worksheet.getCell(`Q1`).value !== 'Dia 2 Cut' ||
        worksheet.getCell(`R1`).value !== 'Dia 2 Polish' ||
        worksheet.getCell(`S1`).value !== 'Dia 2 Symmetry' ||
        worksheet.getCell(`T1`).value !== 'Dia 2 Certificate' ||
        worksheet.getCell(`U1`).value !== 'Dia 3 Carat' ||
        worksheet.getCell(`V1`).value !== 'Dia 3 Color' ||
        worksheet.getCell(`W1`).value !== 'Dia 3 Clarity' ||
        worksheet.getCell(`X1`).value !== 'Dia 3 Shape' ||
        worksheet.getCell(`Y1`).value !== 'Dia 3 Cut' ||
        worksheet.getCell(`Z1`).value !== 'Dia 3 Polish' ||
        worksheet.getCell(`AA1`).value !== 'Dia 3 Symmetry' ||
        worksheet.getCell(`AB1`).value !== 'Dia 3 Certificate' ||
        worksheet.getCell(`AC1`).value !== 'Purchase Price (RSD/HUF)' ||
        worksheet.getCell(`AD1`).value !== 'Retail Price' ||
        worksheet.getCell(`AE1`).value !== 'Serial Number' ||
        worksheet.getCell(`AF1`).value !== 'PGP' ||
        worksheet.getCell(`AG1`).value !== 'Location' ||
        worksheet.getCell(`AH1`).value !== 'Comment' ||
        worksheet.getCell(`AI1`).value !== 'Stock Date' ||
        worksheet.getCell(`AJ1`).value !== 'Invoice Number' ||
        worksheet.getCell(`AK1`).value !== 'Invoice Date'
      ) throw new Error(error.INVALID_COLUMN_NAME);
      break;
  }

  // For each row in the 'Tppack' table
  for (let i = 0; i < worksheet.actualRowCount - 1; i++) {
    // Basic info properties
    let rmc = '';
    let photos = [];
    let collection = '';
    let saleReference = '';
    let materialDescription = '';
    let jewelryType = '';
    let size = '';
    let weight = '';
    let allStonesWeight = '';
    let stonesQty = '';
    let brilliants = '';
    let diaGia = '';
    let purchasePrice = '';
    let materials = [];
    let stones = [];
    let diamonds = [];

    // Boutique properties
    let boutiques = [];
    let priceLocal = null;
    let priceHistoryRs = [];
    let serialNumbers = [];

    // Serial number properties
    let serialNumber = '';
    let pgpReference = '';
    let location = '';
    let comment = '';
    let stockDate = null;
    let invoiceNumber = '';
    let invoiceDate = null;
    let exGenevaPrice = null;
    let purchasePriceLocal = null;

    // New product and activity
    let newProduct = null;
    let updateProduct = null;
    let newActivity = null;
    let activityDate = null;
    let logComment = '';

    // Product to be found in DB
    let product = null;

    switch (brand) {
      // PANERAI
      case 'Panerai':
        // Get added watch features
        saleReference = worksheet.getCell(`A${i + 2}`).value.toString();
        serialNumber = worksheet.getCell(`B${i + 2}`).value;
        pgpReference = worksheet.getCell(`C${i + 2}`).value;
        location = worksheet.getCell(`D${i + 2}`).value;
        comment = worksheet.getCell(`E${i + 2}`).value ? worksheet.getCell(`E${i + 2}`).value : '';
        stockDate = worksheet.getCell(`F${i + 2}`).value;
        invoiceDate = worksheet.getCell(`G${i + 2}`).value;
        purchasePriceLocal = worksheet.getCell(`H${i + 2}`).value;

        // Check required properties
        if (!saleReference || !serialNumber || !pgpReference) throw new Error(error.MISSING_PARAMETERS);
        if (!location) throw new Error(error.MISSING_LOCATION);

        // Check if dates are valid
        if (stockDate && isNaN(Date.parse(stockDate))) throw new Error(error.INVALID_STOCK_DATE);
        if (invoiceDate && isNaN(Date.parse(invoiceDate))) throw new Error(error.INVALID_INVOICE_DATE);

        // Find product in DB with the sent 'rmc'
        product = existingProducts.find(obj => obj.basicInfo.saleReference === saleReference);

        // Check if product was found
        if (!product) throw new Error(error.NOT_FOUND);

        // 1. Update product -> add watch to stock
        updateProduct = Product.updateOne(
          {
            'basicInfo.saleReference': saleReference,
            boutiques: { $elemMatch: { store } }
          },
          {
            $addToSet: {
              'boutiques.$.serialNumbers': {
                number: serialNumber,
                pgpReference,
                location,
                status: `New stock`,
                comment,
                stockDate: stockDate ? stockDate : new Date(),
                invoiceDate,
                purchasePriceLocal
              }
            },
            $inc: { 'boutiques.$.quantity': 1 }
          },
        );

        // Push updated product to 'newWatches' array
        addToStocks.push(updateProduct);

        // 2. Create new activity
        activityDate = stockDate ? new Date(stockDate) : new Date();
        logComment = `Watch added to stock`;
        newActivity = createActivity('Product', userId, null, logComment, null, product._id, serialNumber, null, activityDate);

        // Push created activity to 'newActivities' array
        newActivities.push(newActivity.save());
        break;

      // SWISSKUBIK
      case 'SwissKubik':
        // Get added watch features
        saleReference = worksheet.getCell(`A${i + 2}`).value.toString();
        serialNumber = worksheet.getCell(`B${i + 2}`).value;
        pgpReference = worksheet.getCell(`C${i + 2}`).value;
        location = worksheet.getCell(`D${i + 2}`).value;
        comment = worksheet.getCell(`E${i + 2}`).value ? worksheet.getCell(`E${i + 2}`).value : '';
        stockDate = worksheet.getCell(`F${i + 2}`).value;
        invoiceNumber = worksheet.getCell(`G${i + 2}`).value;
        invoiceDate = worksheet.getCell(`H${i + 2}`).value;
        exGenevaPrice = worksheet.getCell(`I${i + 2}`).value;
        purchasePriceLocal = worksheet.getCell(`J${i + 2}`).value;

        // Check required properties
        if (!saleReference || !serialNumber || !pgpReference) throw new Error(error.MISSING_PARAMETERS);
        if (!location) throw new Error(error.MISSING_LOCATION);
        if (!invoiceDate) throw new Error(error.MISSING_INVOICE_DATE);
        if (!exGenevaPrice) throw new Error(error.MISSING_EX_GENEVA_PRICE);

        // Check if dates are valid
        if (stockDate && isNaN(Date.parse(stockDate))) throw new Error(error.INVALID_STOCK_DATE);
        if (invoiceDate && isNaN(Date.parse(invoiceDate))) throw new Error(error.INVALID_INVOICE_DATE);

        // Find product in DB with the sent 'rmc'
        product = existingProducts.find(obj => obj.basicInfo.saleReference === saleReference);

        // Check if product was found
        if (!product) throw new Error(error.NOT_FOUND);

        // 1. Update product -> add watch to stock
        updateProduct = Product.updateOne(
          {
            'basicInfo.saleReference': saleReference,
            boutiques: { $elemMatch: { store } }
          },
          {
            $addToSet: {
              'boutiques.$.serialNumbers': {
                number: serialNumber,
                pgpReference,
                location,
                status: `New stock`,
                comment,
                stockDate: stockDate ? stockDate : new Date(),
                invoiceNumber,
                invoiceDate,
                exGenevaPrice,
                purchasePriceLocal
              }
            },
            $inc: { 'boutiques.$.quantity': 1 }
          },
        );

        // Push updated product to 'newWatches' array
        addToStocks.push(updateProduct);

        // 2. Create new activity
        activityDate = stockDate ? new Date(stockDate) : new Date();
        logComment = `Product added to stock`;
        newActivity = createActivity('Product', userId, null, logComment, null, product._id, serialNumber, null, activityDate);

        // Push created activity to 'newActivities' array
        newActivities.push(newActivity.save());
        break;

      // RUBBER B
      case 'Rubber B':
        // Get added watch features
        rmc = worksheet.getCell(`A${i + 2}`).value.toString();
        serialNumber = worksheet.getCell(`B${i + 2}`).value;
        pgpReference = worksheet.getCell(`C${i + 2}`).value;
        location = worksheet.getCell(`D${i + 2}`).value;
        adjustedSize = worksheet.getCell(`E${i + 2}`).value ? worksheet.getCell(`E${i + 2}`).value.toUpperCase() : '';
        comment = worksheet.getCell(`F${i + 2}`).value ? worksheet.getCell(`F${i + 2}`).value : '';
        stockDate = worksheet.getCell(`G${i + 2}`).value;
        invoiceDate = worksheet.getCell(`H${i + 2}`).value;
        purchasePriceLocal = worksheet.getCell(`I${i + 2}`).value;

        // Check required properties
        if (!rmc || !serialNumber || !pgpReference) throw new Error(error.MISSING_PARAMETERS);
        if (!location) throw new Error(error.MISSING_LOCATION);

        // Validation
        if (stockDate && isNaN(Date.parse(stockDate))) throw new Error(error.INVALID_STOCK_DATE);
        if (invoiceDate && isNaN(Date.parse(invoiceDate))) throw new Error(error.INVALID_INVOICE_DATE);
        if (adjustedSize && !['XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL', 'XXXXL'].includes(adjustedSize)) throw new Error(error.INVALID_SIZE);

        // Find product in DB with the sent 'rmc'
        product = existingProducts.find(obj => obj.basicInfo.rmc === rmc);

        // Check if product was found
        if (!product) throw new Error(error.NOT_FOUND);

        // 1. Update product -> add watch to stock
        updateProduct = Product.updateOne(
          {
            'basicInfo.rmc': rmc,
            boutiques: { $elemMatch: { store } }
          },
          {
            $addToSet: {
              'boutiques.$.serialNumbers': {
                number: serialNumber,
                pgpReference,
                location,
                status: `New stock`,
                adjustedSize,
                comment,
                stockDate: stockDate ? stockDate : new Date(),
                invoiceDate,
                purchasePriceLocal
              }
            },
            $inc: { 'boutiques.$.quantity': 1 }
          },
        );

        // Push updated product to 'newWatches' array
        addToStocks.push(updateProduct);

        // 2. Create new activity
        activityDate = stockDate ? new Date(stockDate) : new Date();
        logComment = `Product added to stock`;
        newActivity = createActivity('Product', userId, null, logComment, null, product._id, serialNumber, null, activityDate);

        // Push created activity to 'newActivities' array
        newActivities.push(newActivity.save());
        break;

      // MESSIKA
      case 'Messika':
        // Get added watch features
        collection = worksheet.getCell(`A${i + 2}`).value.toUpperCase();
        saleReference = worksheet.getCell(`B${i + 2}`).value;
        rmc = saleReference.slice(0,7).split('-').reverse().join('');
        photos = [
          `${rmc}.jpg`,
          `${rmc}_p.jpg`
        ];
        materialDescription = worksheet.getCell(`C${i + 2}`).value;
        jewelryType = worksheet.getCell(`D${i + 2}`).value ? worksheet.getCell(`D${i + 2}`).value.toLowerCase() : '';
        // Set 'materials' based on RMC number (if it is missing)
        switch (rmc[0]) {
          case 'P':
            materials = ['Pink gold'];
            break;
          case 'Y':
            materials = ['Yellow gold'];
            break;
          case 'W':
            materials = ['White gold'];
            break;
          case 'N':
            materials = ['Natural Titanium'];
            break;
          case 'W':
            materials = ['Graphite Titanium'];
            break;
        }
        size = worksheet.getCell(`E${i + 2}`).value ? worksheet.getCell(`E${i + 2}`).value : '';
        weight = worksheet.getCell(`F${i + 2}`).value;
        allStonesWeight = worksheet.getCell(`G${i + 2}`).value;
        stonesQty = worksheet.getCell(`H${i + 2}`).value;
        brilliants = worksheet.getCell(`I${i + 2}`).value ? worksheet.getCell(`I${i + 2}`).value : '';
        purchasePrice = worksheet.getCell(`J${i + 2}`).value;

        purchasePriceLocal = worksheet.getCell(`K${i + 2}`).value;

        retailRsEUR = worksheet.getCell(`L${i + 2}`).value;
        priceLocal = Math.ceil(retailRsEUR * rateRSDJewelry / 1000) * 1000;   // round up to the nearest thousand

        serialNumber = worksheet.getCell(`M${i + 2}`).value;
        pgpReference = worksheet.getCell(`N${i + 2}`).value;
        location = worksheet.getCell(`O${i + 2}`).value;
        comment = worksheet.getCell(`P${i + 2}`).value ? worksheet.getCell(`P${i + 2}`).value : '';
        stockDate = worksheet.getCell(`Q${i + 2}`).value ? worksheet.getCell(`Q${i + 2}`).value : new Date();
        invoiceNumber = worksheet.getCell(`R${i + 2}`).value ? worksheet.getCell(`R${i + 2}`).value : '';
        invoiceDate = worksheet.getCell(`S${i + 2}`).value;

        // Check required properties
        if (!saleReference || !collection || !serialNumber || !pgpReference) throw new Error(error.MISSING_PARAMETERS);
        if (!retailRsEUR) throw new Error(error.MISSING_PRICE);
        if (!jewelryType) throw new Error(error.MISSING_JEWELRY_TYPE);
        if (!location) throw new Error(error.MISSING_LOCATION);
        if (!purchasePrice) throw new Error(error.MISSING_PURCHASE_PRICE);

        // Check if dates are valid
        if (stockDate && isNaN(Date.parse(stockDate))) throw new Error(error.INVALID_STOCK_DATE);
        if (invoiceDate && isNaN(Date.parse(invoiceDate))) throw new Error(error.INVALID_INVOICE_DATE);

        // Validate 'jewelryType'
        if (jewelryType && !jewelryTypes.includes(jewelryType)) throw new Error(error.INVALID_JEWELRY_TYPE);

        // Create price history arrays
        priceHistoryRs = [{ date: stockDate, price: retailRsEUR, VAT: vatRs, priceLocal }];

        // Create 'boutiques' array
        serialNumbers = [
          {
            number: serialNumber,
            pgpReference,
            status: `New stock`,
            location,
            stockDate,
            comment,
            invoiceDate,
            invoiceNumber,
            purchasePriceLocal
          }
        ];

        // For each existing store create object containing: store ID, store name, respective price and VAT percent, and push the object into 'boutiques' array
        for (const store of existingStores) {
          if (store.name === 'Belgrade') boutiques.push({
            store: store._id,
            storeName: store.name,
            price: retailRsEUR,
            priceLocal,
            VATpercent: vatRs,
            priceHistory: priceHistoryRs,
            quantity: 1,
            serialNumbers
          });
        }

        // 1. Create new product
        newProduct = new Product({
          status: 'new',
          brand,
          boutiques,
          basicInfo: {
            rmc,
            photos,
            collection,
            saleReference,
            materialDescription,
            jewelryType,
            size,
            weight,
            stonesQty,
            allStonesWeight,
            brilliants,
            // diaGia,
            purchasePrice,
            materials,
            // stones,
            // diamonds
          }
        });

        // Push created product to 'newProducts' array
        newProducts.push(newProduct);

        // 2. Create new activity
        activityDate = stockDate ? new Date(stockDate) : new Date();
        logComment = `Product added to stock`;
        newActivity = createActivity('Product', userId, null, logComment, null, newProduct._id, serialNumber, null, activityDate);

        // Push created activity to 'newActivities' array
        newActivities.push(newActivity.save());
        break;

      // ROBERTO COIN
      case 'Roberto Coin':
        // Get added watch features
        collection = worksheet.getCell(`A${i + 2}`).value.toUpperCase();
        saleReference = worksheet.getCell(`B${i + 2}`).value;
        rmc = saleReference;
        photos = [
          `${rmc}.jpg`,
          `${rmc}_p.jpg`
        ];
        jewelryType = worksheet.getCell(`C${i + 2}`).value ? worksheet.getCell(`C${i + 2}`).value.toLowerCase() : '';
        size = worksheet.getCell(`D${i + 2}`).value ? worksheet.getCell(`D${i + 2}`).value : '';
        // Set 'materials' based on RMC number (if it is missing)
        switch (worksheet.getCell(`E${i + 2}`).value) {
          case 'R':
            materials = ['Rose gold'];
            break;
          case 'Y':
            materials = ['Yellow gold'];
            break;
          case 'W':
            materials = ['White gold'];
            break;
          case 'B':
            materials = ['Black gold'];
            break;
          case 'P':
            materials = ['Pink gold'];
            break;
          case 'RW':
            materials = ['Rose gold', 'White gold'];
            break;
          case 'RB':
            materials = ['Rose gold', 'Black gold'];
            break;
          case 'WB':
            materials = ['White gold', 'Black gold'];
            break;
          case 'RBW':
            materials = ['Rose gold', 'Black gold', 'White gold'];
            break;
          case 'RWB':
            materials = ['Rose gold', 'White gold', 'Black gold'];
            break;
        }
        weight = worksheet.getCell(`F${i + 2}`).value ? worksheet.getCell(`F${i + 2}`).value : '';
        diaGia = worksheet.getCell(`G${i + 2}`).value;
        diaQty = worksheet.getCell(`H${i + 2}`).value;
        diaCarat = worksheet.getCell(`I${i + 2}`).value;
        rubyQty = worksheet.getCell(`J${i + 2}`).value;
        rubyCarat = worksheet.getCell(`K${i + 2}`).value;
        brilliants = worksheet.getCell(`L${i + 2}`).value ? worksheet.getCell(`L${i + 2}`).value : '';
        purchasePrice = worksheet.getCell(`M${i + 2}`).value;

        purchasePriceLocal = worksheet.getCell(`N${i + 2}`).value;

        retailRsEUR = worksheet.getCell(`O${i + 2}`).value;
        priceLocal = Math.ceil(retailRsEUR * rateRSDJewelry / 1000) * 1000;   // round up to the nearest thousand

        serialNumber = worksheet.getCell(`P${i + 2}`).value;
        pgpReference = worksheet.getCell(`Q${i + 2}`).value;
        location = worksheet.getCell(`R${i + 2}`).value;
        comment = worksheet.getCell(`S${i + 2}`).value ? worksheet.getCell(`S${i + 2}`).value : '';
        stockDate = worksheet.getCell(`T${i + 2}`).value ? worksheet.getCell(`T${i + 2}`).value : new Date();
        invoiceNumber = worksheet.getCell(`U${i + 2}`).value ? worksheet.getCell(`U${i + 2}`).value : '';
        invoiceDate = worksheet.getCell(`V${i + 2}`).value;

        // Check required properties
        if (!saleReference || !collection || !serialNumber || !pgpReference) throw new Error(error.MISSING_PARAMETERS);
        if (!retailRsEUR) throw new Error(error.MISSING_PRICE);
        if (!jewelryType) throw new Error(error.MISSING_JEWELRY_TYPE);
        if (!location) throw new Error(error.MISSING_LOCATION);
        if (!purchasePrice) throw new Error(error.MISSING_PURCHASE_PRICE);

        // Check if dates are valid
        if (stockDate && isNaN(Date.parse(stockDate))) throw new Error(error.INVALID_STOCK_DATE);
        if (invoiceDate && isNaN(Date.parse(invoiceDate))) throw new Error(error.INVALID_INVOICE_DATE);

        // Validate 'jewelryType'
        if (jewelryType && !jewelryTypes.includes(jewelryType)) throw new Error(error.INVALID_JEWELRY_TYPE);

        // Create price history arrays
        priceHistoryRs = [{ date: stockDate, price: retailRsEUR, VAT: vatRs, priceLocal }];

        // Create 'boutiques' array
        serialNumbers = [
          {
            number: serialNumber,
            pgpReference,
            status: `New stock`,
            location,
            stockDate,
            comment,
            invoiceDate,
            invoiceNumber,
            purchasePriceLocal
          }
        ];

        // For each existing store create object containing: store ID, store name, respective price and VAT percent, and push the object into 'boutiques' array
        for (const store of existingStores) {
          if (store.name === 'Belgrade') boutiques.push({
            store: store._id,
            storeName: store.name,
            price: retailRsEUR,
            priceLocal,
            VATpercent: vatRs,
            priceHistory: priceHistoryRs,
            quantity: 1,
            serialNumbers
          });
        }

        // Create 'stones' array
        if (rubyQty || rubyCarat) {
          stones = [
            {
              type: 'Ruby',
              quantity: rubyQty,
              stoneTypeWeight: rubyCarat ? rubyCarat : ''
            }
          ];
        }

        // Validate 'stones' array
        if (stones && stones.length && !stones.every(el => stoneTypes.includes(el.type))) throw new Error(error.INVALID_STONE_TYPE);

        // // In case of additional type of stone
        // if (stoneType && stoneType !== '') stones.push({
        //   type: stoneType,
        //   quantity: stoneQty,
        //   stoneTypeWeight: stoneCarat
        // });

        // Create 'diamonds' array
        if (diaQty || diaCarat) {
          diamonds = [
            {
              quantity: diaQty,
              carat: diaCarat ? diaCarat : '',
              // giaReports: certificates
            }
          ];
        }

        // 1. Create new product
        newProduct = new Product({
          status: 'new',
          brand,
          boutiques,
          basicInfo: {
            rmc,
            photos,
            collection,
            saleReference,
            // materialDescription,
            jewelryType,
            size,
            weight,
            stonesQty: diaQty + rubyQty,
            allStonesWeight: diaCarat && rubyCarat ? diaCarat + rubyCarat : diaCarat && !rubyCarat ? diaCarat : !diaCarat && rubyCarat ? rubyCarat : '',
            brilliants,
            diaGia,
            purchasePrice,
            materials,
            stones,
            diamonds
          }
        });

        // Push created product to 'newProducts' array
        newProducts.push(newProduct);

        // 2. Create new activity
        activityDate = stockDate ? new Date(stockDate) : new Date();
        logComment = `Product added to stock`;
        newActivity = createActivity('Product', userId, null, logComment, null, newProduct._id, serialNumber, null, activityDate);

        // Push created activity to 'newActivities' array
        newActivities.push(newActivity.save());
        break;

      // PETROVIC DIAMONDS
      case 'Petrovic Diamonds':
        // Get added watch features
        saleReference = worksheet.getCell(`A${i + 2}`).value;
        rmc = saleReference;
        photos = [
          `${rmc}.jpg`,
          `${rmc}_second.jpg`,
          `${rmc}_third.jpg`
        ];
        jewelryType = worksheet.getCell(`B${i + 2}`).value ? worksheet.getCell(`B${i + 2}`).value.toLowerCase() : '';
        size = worksheet.getCell(`C${i + 2}`).value ? worksheet.getCell(`C${i + 2}`).value : '';
        weight = worksheet.getCell(`D${i + 2}`).value ? worksheet.getCell(`D${i + 2}`).value : '';
        dia1Carat = worksheet.getCell(`E${i + 2}`).value ? worksheet.getCell(`E${i + 2}`).value : '';
        dia1Color = worksheet.getCell(`F${i + 2}`).value ? worksheet.getCell(`F${i + 2}`).value.toUpperCase() : '';
        dia1Clarity = worksheet.getCell(`G${i + 2}`).value ? worksheet.getCell(`G${i + 2}`).value.toUpperCase() : '';
        dia1Shape = worksheet.getCell(`H${i + 2}`).value ? worksheet.getCell(`H${i + 2}`).value.toLowerCase() : '';
        dia1Cut = worksheet.getCell(`I${i + 2}`).value ? worksheet.getCell(`I${i + 2}`).value.toLowerCase() : '';
        dia1Polish = worksheet.getCell(`J${i + 2}`).value ? worksheet.getCell(`J${i + 2}`).value : '';
        dia1Symmetry = worksheet.getCell(`K${i + 2}`).value ? worksheet.getCell(`K${i + 2}`).value : '';
        dia1Certificate = worksheet.getCell(`L${i + 2}`).value ? worksheet.getCell(`L${i + 2}`).value : '';
        dia2Carat = worksheet.getCell(`M${i + 2}`).value ? worksheet.getCell(`M${i + 2}`).value : '';
        dia2Color = worksheet.getCell(`N${i + 2}`).value ? worksheet.getCell(`N${i + 2}`).value.toUpperCase() : '';
        dia2Clarity = worksheet.getCell(`O${i + 2}`).value ? worksheet.getCell(`O${i + 2}`).value.toUpperCase() : '';
        dia2Shape = worksheet.getCell(`P${i + 2}`).value ? worksheet.getCell(`P${i + 2}`).value.toLowerCase() : '';
        dia2Cut = worksheet.getCell(`Q${i + 2}`).value ? worksheet.getCell(`Q${i + 2}`).value.toLowerCase() : '';
        dia2Polish = worksheet.getCell(`R${i + 2}`).value ? worksheet.getCell(`R${i + 2}`).value : '';
        dia2Symmetry = worksheet.getCell(`S${i + 2}`).value ? worksheet.getCell(`S${i + 2}`).value : '';
        dia2Certificate = worksheet.getCell(`T${i + 2}`).value ? worksheet.getCell(`T${i + 2}`).value : '';
        dia3Carat = worksheet.getCell(`U${i + 2}`).value ? worksheet.getCell(`U${i + 2}`).value : '';
        dia3Color = worksheet.getCell(`V${i + 2}`).value ? worksheet.getCell(`V${i + 2}`).value.toUpperCase() : '';
        dia3Clarity = worksheet.getCell(`W${i + 2}`).value ? worksheet.getCell(`W${i + 2}`).value.toUpperCase() : '';
        dia3Shape = worksheet.getCell(`X${i + 2}`).value ? worksheet.getCell(`X${i + 2}`).value.toLowerCase() : '';
        dia3Cut = worksheet.getCell(`Y${i + 2}`).value ? worksheet.getCell(`Y${i + 2}`).value.toLowerCase() : '';
        dia3Polish = worksheet.getCell(`Z${i + 2}`).value ? worksheet.getCell(`Z${i + 2}`).value : '';
        dia3Symmetry = worksheet.getCell(`AA${i + 2}`).value ? worksheet.getCell(`AA${i + 2}`).value : '';
        dia3Certificate = worksheet.getCell(`AB${i + 2}`).value ? worksheet.getCell(`AB${i + 2}`).value : '';

        purchasePriceLocal = worksheet.getCell(`AC${i + 2}`).value;

        retailRsEUR = worksheet.getCell(`AD${i + 2}`).value;
        priceLocal = Math.ceil(retailRsEUR * rateRSDJewelry / 1000) * 1000;   // round up to the nearest thousand

        serialNumber = worksheet.getCell(`AE${i + 2}`).value;
        pgpReference = worksheet.getCell(`AF${i + 2}`).value;
        location = worksheet.getCell(`AG${i + 2}`).value;
        comment = worksheet.getCell(`AH${i + 2}`).value ? worksheet.getCell(`AH${i + 2}`).value : '';
        stockDate = worksheet.getCell(`AI${i + 2}`).value ? worksheet.getCell(`AI${i + 2}`).value : new Date();
        invoiceNumber = worksheet.getCell(`AJ${i + 2}`).value ? worksheet.getCell(`AJ${i + 2}`).value : '';
        invoiceDate = worksheet.getCell(`AK${i + 2}`).value;

        // Check required properties
        if (!saleReference || !serialNumber || !pgpReference) throw new Error(error.MISSING_PARAMETERS);
        if (!retailRsEUR) throw new Error(error.MISSING_PRICE);
        if (!jewelryType) throw new Error(error.MISSING_JEWELRY_TYPE);
        if (!location) throw new Error(error.MISSING_LOCATION);

        // Check if dates are valid
        if (stockDate && isNaN(Date.parse(stockDate))) throw new Error(error.INVALID_STOCK_DATE);
        if (invoiceDate && isNaN(Date.parse(invoiceDate))) throw new Error(error.INVALID_INVOICE_DATE);

        // Set 'diamonds' array
        let dia1 = null;
        if (dia1Carat || dia1Color || dia1Clarity || dia1Shape || dia1Cut || dia1Polish || dia1Symmetry || dia1Certificate) dia1 = {
          quantity: 1,
          carat: dia1Carat,
          color: dia1Color,
          clarity: dia1Clarity,
          shape: dia1Shape,
          cut: dia1Cut,
          polish: dia1Polish,
          symmetry: dia1Symmetry,
          giaReports: dia1Certificate,
          giaReportsUrls: `http://192.168.2.50:9000/minio/download/rolex/${dia1Certificate}.pdf?token=`
        };

        let dia2 = null;
        if (dia2Carat || dia2Color || dia2Clarity || dia2Shape || dia2Cut || dia2Polish || dia2Symmetry || dia2Certificate) dia2 = {
          quantity: 1,
          carat: dia2Carat,
          color: dia2Color,
          clarity: dia2Clarity,
          shape: dia2Shape,
          cut: dia2Cut,
          polish: dia2Polish,
          symmetry: dia2Symmetry,
          giaReports: dia2Certificate,
          giaReportsUrls: `http://192.168.2.50:9000/minio/download/rolex/${dia2Certificate}.pdf?token=`
        };

        let dia3 = null;
        if (dia3Carat || dia3Color || dia3Clarity || dia3Shape || dia3Cut || dia3Polish || dia3Symmetry ||  dia3Certificate) dia3 = {
          quantity: 1,
          carat: dia3Carat,
          color: dia3Color,
          clarity: dia3Clarity,
          shape: dia3Shape,
          cut: dia3Cut,
          polish: dia3Polish,
          symmetry: dia3Symmetry,
          giaReports: dia3Certificate,
          giaReportsUrls: `${environments.MINIO_DOWNLOAD_URL}${dia3Certificate}.pdf?token=`
        };

        if (dia1) diamonds.push(dia1);
        if (dia2) diamonds.push(dia2);
        if (dia3) diamonds.push(dia3);

        // Validate 'diamonds' array
        if (diamonds && diamonds.length && !diamonds.every(el => colors.includes(el.color))) throw new Error(error.INVALID_COLOR);
        if (diamonds && diamonds.length && !diamonds.every(el => clarities.includes(el.clarity))) throw new Error(error.INVALID_CLARITY);
        if (diamonds && diamonds.length && !diamonds.every(el => shapes.includes(el.shape))) throw new Error(error.INVALID_SHAPE);
        if (diamonds && diamonds.length && !diamonds.every(el => cuts.includes(el.cut))) throw new Error(error.INVALID_CUT);

        allStonesWeight = Math.round(diamonds.reduce((total, { carat }) => total + Number(carat), 0) * 10000)/10000;

        // Create price history arrays
        priceHistoryRs = [{ date: stockDate, price: retailRsEUR, VAT: vatRs, priceLocal }];

        // Create 'boutiques' array
        serialNumbers = [
          {
            number: serialNumber,
            pgpReference,
            location,
            status: `New stock`,
            comment,
            stockDate,
            invoiceDate,
            invoiceNumber,
            purchasePriceLocal
          }
        ];

        // For each existing store create object containing: store ID, store name, respective price and VAT percent, and push the object into 'boutiques' array
        for (const store of existingStores) {
          if (store.name === 'Belgrade') boutiques.push({
            store: store._id,
            storeName: store.name,
            price: retailRsEUR,
            priceLocal,
            VATpercent: vatRs,
            priceHistory: priceHistoryRs,
            quantity: 1,
            serialNumbers
          });
        }

        // 1. Create new product
        newProduct = new Product({
          status: 'new',
          brand,
          boutiques,
          basicInfo: {
            rmc,
            photos,
            // collection
            saleReference,
            // materialDescription,
            jewelryType,
            size,
            weight,
            stonesQty: diamonds.length,
            allStonesWeight,
            // brilliants,
            // diaGia,
            // purchasePrice,
            // materials,
            // stones,
            diamonds
          }
        });

        // Push created product to 'newProducts' array
        newProducts.push(newProduct);

        // 2. Create new activity
        activityDate = stockDate ? new Date(stockDate) : new Date();
        logComment = `Product added to stock`;
        newActivity = createActivity('Product', userId, null, logComment, null, newProduct._id, serialNumber, null, activityDate);

        // Push created activity to 'newActivities' array
        newActivities.push(newActivity.save());
        break;
    }
  }

  // Execute
  await Promise.all([
    ...addToStocks,
    Product.insertMany(newProducts),
    ...newActivities
  ]);

  return res.status(201).send({
    message: 'Successfully imported new stock',
  });
};

/**
 * @api {get} /product/stock Get all soon in stock watches
 * @apiVersion 1.0.0
 * @apiName soonInStockWatches
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} [dial] Filter by Dial
 * @apiParam (query) {String} [store] Filter by Store ID
 * @apiParam (query) {String} [searchRmc] Search watches by RMC
 * @apiParam (query) {String} [searchSerialNumber] Search watches by Serial Number
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned all soon in stock watches",
   "count": 3,
   "results": [
     {
       "_id": "5f8d3788cea46482dcd6c4e2",
       "rmc": "M116769TBRJ-0002",
       "serialNumber": "G8L14630",
       "dial": "PAVED W",
       "store": {
         "_id": "5f8d3788cea46482dcd6c4e1",
         "name": "Belgrade"
       },
       "__v": 0,
       "createdAt": "2020-10-19T06:51:52.701Z",
       "updatedAt": "2020-10-19T06:51:52.701Z"
     },
     ...
   ]
 }
 * @apiUse InvalidValue
 */
module.exports.soonInStockWatches = async (req, res) => {
  const { dial, store, searchRmc, searchSerialNumber } = req.query;

  // Creat query object
  let query = {};

  // Check if 'dial' filter has been sent
  if (dial) query.dial = dial;

  // Check if 'store' filter has been sent
  if (store) query.store = store;

  // Search via RMC number filtered products
  if (searchRmc) query.rmc = new RegExp(`.*${searchRmc}.*`, 'i');

  // Search via serial number filtered products
  if (searchSerialNumber) query.serialNumber = new RegExp(`.*${searchSerialNumber}.*`, 'i');

  // Get list of soon in stock products and count
  let [listOfWatches, count] = await Promise.all([
    SoonInStock.find(query)
      .populate('store', 'name')
      .sort('pgpReference')
      .lean(),
    SoonInStock.countDocuments(query).lean(),
  ]);

  return res.status(200).send({
    message: 'Successfully returned all soon in stock watches',
    count,
    results: listOfWatches,
  });
};

/**
 * @api {get} /product/soonInStock Get soon in stock product details
 * @apiVersion 1.0.0
 * @apiName soonInStockDetails
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} serialNumber Soon in stock serial number
 * @apiParam (query) {String} storeId Store ID
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned soon in stock product details",
   "results": {
     "_id": "5fe9c5f46eb886c457d49bb9",
     "reservationTime": "2021-03-30T22:00:00.000Z",
     "reservedFor": {
       "_id": "60425091e17126e9f5bc630e",
       "fullName": "Milos Vladimir Pavlovic"
     },
     "soonInStock": true,
     "warrantyConfirmed": true,
     "rmc": "iusto",
     "product": {
       "_id": "5fe9c5f46eb886c457d49bb3",
       "basicInfo": {
         "photos": [
           "http://antonette.com"
         ],
         "materials": [],
         "rmc": "iusto",
         "collection": "TUDOR ARCHEO",
         "productLine": "iste",
         "saleReference": "tempore",
         "materialDescription": "id",
         "dial": "tenetur",
         "bracelet": "ipsam",
         "box": "enim",
         "exGeneveCHF": 51110,
         "diameter": "80546",
         "stones": [],
         "diamonds": []
       },
       "brand": "Tudor"
     },
     "serialNumber": "4694",
     "dial": "doloremque",
     "pgpReference": "et",
     "store": {
       "_id": "5fe9c5f46eb886c457d49bb8",
       "name": "Joannyborough"
     },
     "origin": "West Alphonsoburgh",
     "location": "program",
     "status": "Stock",
     "stockDate": "2020-12-28T10:49:04.366Z",
     "card": "vel",
     "comment": "nihil",
     "soldToParty": "iure",
     "shipToParty": "hic",
     "billToParty": "quod",
     "packingNumber": "at",
     "invoiceDate": "2020-12-27T16:47:09.640Z",
     "invoiceNumber": "non",
     "boxCode": "hic",
     "exGenevaPrice": 3,
     "sectorId": "fuga",
     "sector": "iusto",
     "boitId": "iste",
     "boitRef": "provident",
     "caspId": "harum",
     "cadranDesc": "delectus",
     "ldisId": "corporis",
     "ldisq": "quas",
     "brspId": "eos",
     "brspRef": "nemo",
     "prixSuisse": 8,
     "umIntern": "qui",
     "umExtern": "aut",
     "rfid": "odit",
     "packingType": "recusandae",
     "__v": 0,
     "createdAt": "2020-12-28T11:48:04.213Z",
     "updatedAt": "2020-12-28T11:48:04.213Z"
   },
   "activities": [
     {
       "_id": "607727cc0d06b335a84eb899",
       "type": "Product",
       "client": null,
       "comment": "Changed location from 'Expected on 15.05.' to 'Expected on 15.03.'. Changed status from 'Reserved' to 'Active'. Comment Changed reservation from 'Endre Sic' to 'Martin McKay'. Changed reservation time from '15/04/2021' to '15/04/2021'.",
       "wishlist": null,
       "product": "603caa9776077912220b384c",
       "serialNumber": "M1_99251",
       "createdAt": "2021-04-14T17:35:08.747Z",
       "__v": 0
     }
   ]
 }
 * @apiUse NotFound
 * @apiUse MissingParamsError
 */
module.exports.soonInStockDetails = async (req, res) => {
  const { serialNumber, storeId } = req.query;

  // Check if required data has been sent
  if (!serialNumber || !storeId) throw new Error(error.MISSING_PARAMETERS);

  // Check if storeId is valid ObjectId
  if (!isValidId(storeId)) throw new Error(error.INVALID_VALUE);

  // Find soon in stock product and all activities related to the sent 'serialNumber'
  const [soonInStockProduct, activities] = await Promise.all([
    SoonInStock.findOne({ serialNumber, store: storeId }).populate('reservedFor', 'fullName').populate('store', '_id name').populate('product', '_id basicInfo status brand').lean(),
    Activity.find({ type: 'Product', serialNumber }).populate('user', 'name').sort('-createdAt').lean(),
  ]);

  // Check if soon in stock product is found
  if (!soonInStockProduct) throw new Error(error.NOT_FOUND);

  return res.status(200).send({
    message: 'Successfully returned soon in stock product details',
    results: soonInStockProduct,
    activities,
  });
};

/**
 * @api {patch} /product/soonInStock Edit soon in stock product
 * @apiVersion 1.0.0
 * @apiName editSoonInStock
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} storeId Store ID
 * @apiParam (body) {String} [rmc] Rmc
 * @apiParam (body) {String} [dial] Dial
 * @apiParam (body) {String} serialNumber Serial Number
 * @apiParam (body) {String} [pgpReference] PGP Reference
 * @apiParam (body) {String} [status] Status
 * @apiParam (body) {Date} [stockDate] Stock Date
 * @apiParam (body) {String} [location] Location
 * @apiParam (body) {String} [origin] Origin
 * @apiParam (body) {String} [card] Card
 * @apiParam (body) {Boolean} [warrantyConfirmed] Warranty Confirmed
 * @apiParam (body) {String} [comment] Comment
 * @apiParam (body) {String} [reservedFor] Client ID (watch is reserved for)
 * @apiParam (body) {date} [reservationTime] Date and time when reservation of a watch expires
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 201 OK
 {
   "message": "Successfully updated soon in stock product",
   "results": {
     "_id": "5fe9f3ea4e90c708e163653f",
     "reservationTime": "2021-03-30T22:00:00.000Z",
     "reservedFor": {
       "_id": "60425091e17126e9f5bc630e",
       "fullName": "Milos Vladimir Pavlovic"
     },
     "soonInStock": true,
     "warrantyConfirmed": false,
     "rmc": "dolores",
     "product": {
       "_id": "5fe9f3ea4e90c708e163653a",
       "basicInfo": {
         "photos": [
           "http://aliya.com"
         ],
         "materials": [],
         "rmc": "molestias",
         "collection": "MOVE 10TH",
         "productLine": "inventore",
         "saleReference": "et",
         "materialDescription": "iste",
         "dial": "eum",
         "bracelet": "enim",
         "box": "nisi",
         "exGeneveCHF": 38434,
         "diameter": "1042",
         "stones": [],
         "diamonds": []
       },
       "brand": "Rolex"
     },
     "serialNumber": "expedita",
     "dial": "aut",
     "pgpReference": "quaerat",
     "store": {
       "_id": "5fe9f3ea4e90c708e1636535",
       "name": "South Marshallfort"
     },
     "origin": "nostrum",
     "location": "odit",
     "status": "Not for sale",
     "stockDate": "2020-12-27T16:56:18.965Z",
     "card": "earum",
     "comment": "rem",
     "soldToParty": "aspernatur",
     "shipToParty": "id",
     "billToParty": "rem",
     "packingNumber": "itaque",
     "invoiceDate": "2020-12-27T23:13:28.229Z",
     "invoiceNumber": "iure",
     "boxCode": "sint",
     "exGenevaPrice": 5,
     "sectorId": "consequatur",
     "sector": "enim",
     "boitId": "officia",
     "boitRef": "laboriosam",
     "caspId": "incidunt",
     "cadranDesc": "quam",
     "ldisId": "qui",
     "ldisq": "est",
     "brspId": "ut",
     "brspRef": "quis",
     "prixSuisse": 7,
     "umIntern": "incidunt",
     "umExtern": "qui",
     "rfid": "voluptate",
     "packingType": "dolorem",
     "__v": 0,
     "createdAt": "2020-12-28T15:04:10.236Z",
     "updatedAt": "2020-12-28T15:04:10.252Z"
   }
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
module.exports.editSoonInStock = async (req, res) => {
  const { _id: userId } = req.user;
  const { storeId } = req.query;
  const {
    rmc,
    dial,
    serialNumber,
    pgpReference,
    status,
    stockDate,
    location,
    origin,
    card,
    warrantyConfirmed,
    comment,
    reservedFor,
    reservationTime
  } = req.body;

  if (!serialNumber || !storeId || (!rmc && !dial && !pgpReference && !status && !stockDate && !location && !origin && !card && !warrantyConfirmed && !comment)) throw new Error(error.MISSING_PARAMETERS);

  // Check if storeId is valid ObjectId
  if (!isValidId(storeId)) throw new Error(error.INVALID_VALUE);

  // Check if client exists
  if (reservedFor) {
    // Check if reservedFor is valid ObjectId
    if (!isValidId(reservedFor)) throw new Error(error.INVALID_VALUE);
    const client = await Client.findById(reservedFor).lean();
    if (!client) throw new Error(error.NOT_FOUND);
  }

  const updateSet = {};
  const soonInStock = await SoonInStock.findOne({ serialNumber, store: storeId }).populate('reservedFor', 'fullName').populate('store', '_id name').populate('product', '_id basicInfo status brand').lean();
  if (!soonInStock) throw new Error(error.NOT_FOUND);

  let confirmedReservation = false;
  if (rmc) updateSet.rmc = rmc;
  if (dial) updateSet.dial = dial;
  if (serialNumber) updateSet.serialNumber = serialNumber;
  if (pgpReference) updateSet.pgpReference = pgpReference;
  if (reservedFor || reservedFor === null) updateSet.reservedFor = reservedFor;
  if (reservationTime || reservationTime === '') updateSet.reservationTime = reservationTime;
  if (status) {
    updateSet.status = status;
    if (status === 'Reserved' && soonInStock.status === 'Pre-reserved') {
      confirmedReservation = true;
    } else if (status !== soonInStock.status) {
      updateSet.previousStatus = soonInStock.status;
    }
    if (status !== 'Reserved' && status !== 'Pre-reserved') {
      updateSet.reservedFor = null;
      updateSet.reservationTime = null;
    }
  }
  if (stockDate) updateSet.stockDate = stockDate;
  if (location) updateSet.location = location;
  if (origin) updateSet.origin = origin;
  if (card) updateSet.card = card;
  if (typeof warrantyConfirmed === 'boolean') updateSet.warrantyConfirmed = warrantyConfirmed;
  if (comment) updateSet.comment = comment;

  const results = await SoonInStock.findOneAndUpdate({ serialNumber, store: storeId }, { $set: updateSet }, { new: true }).populate('reservedFor', 'fullName').populate('store', '_id name').populate('product', '_id basicInfo status brand').lean();

  if (!results) throw new Error(error.NOT_FOUND);

  // Create 'changedFields' and 'toExecute' array
  const changedFields = [];
  const toExecute = [];

  // Detect changes in 'location', 'status' and 'comment' -> in order to create history logs
  if (results.location !== soonInStock.location) changedFields.push(`Changed location from '${soonInStock.location}' to '${results.location}'.`);
  if (results.status !== soonInStock.status) changedFields.push(`Changed status from '${soonInStock.status}' to '${results.status}'.`);
  if (results.comment !== '' && (results.comment !== soonInStock.comment)) changedFields.push(comment);

  // Detect 'reservedFor' changes -> create touchpoints for clients journeys
  if (results.reservedFor && !soonInStock.reservedFor) {
    changedFields.push(`Made ${results.status === 'Pre-reserved' ? 'pre-reservation' : 'reservation'} for '${results.reservedFor.fullName}'.`);
  }
  if (!results.reservedFor && soonInStock.reservedFor) {
    changedFields.push(`Cancelled ${results.previousStatus === 'Pre-reserved' ? 'pre-reservation' : 'reservation'} for '${soonInStock.reservedFor.fullName}'.`);
  }
  if (results.reservedFor && soonInStock.reservedFor && (results.reservedFor._id.toString() !== soonInStock.reservedFor._id.toString())) {
    changedFields.push(`Changed ${results.status === 'Pre-reserved' ? 'pre-reservation' : 'reservation'} from '${soonInStock.reservedFor.fullName}' to '${results.reservedFor.fullName}'.`);
  }

  if (results.reservedFor && soonInStock.reservedFor && confirmedReservation) {
    changedFields.push(`Confirmed pre-reservation for '${results.reservedFor.fullName}'.`);
  }

  // Detect 'reservationTime' changes -> in order to create history logs
  if (results.reservationTime && !soonInStock.reservationTime) changedFields.push(`Reservation time set to '${moment(results.reservationTime, 'YYYY-MM-DD').format('DD/MM/YYYY')}'.`);
  if (!results.reservationTime && soonInStock.reservationTime) changedFields.push(`Reservation time previously set to '${moment(soonInStock.reservationTime).format('DD/MM/YYYY')}' has been removed.`);
  if (results.reservationTime && soonInStock.reservationTime && new Date(results.reservationTime).toString() !== soonInStock.reservationTime.toString()) changedFields.push(`Changed reservation time from '${moment(soonInStock.reservationTime).format('DD/MM/YYYY')}' to '${moment(reservationTime, 'YYYY-MM-DD').format('DD/MM/YYYY')}'.`);

  // If there were any changes include those in the log comment
  const logComment = changedFields.length > 0 ? `${changedFields.join(' ')}` : '';

  // Create new activity -> for soonInStock log history
  if (changedFields.length > 0) {
    const newActivity = createActivity('Product', userId, null, logComment, null, results.product._id, results.serialNumber);
    toExecute.push(newActivity.save());
  }

  // Execute
  await Promise.all(toExecute);

  return res.status(200).send({
    message: 'Successfully updated soon in stock product',
    results,
  });
};

/**
 * @api {post} /product/stock Add sooninstocks to stock
 * @apiVersion 1.0.0
 * @apiName addToStock
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (body) {String[]} products Array of soonInStock IDs to be moved to stock
 * @apiParam (body) {String} [location] Internal location of a watch (inside store)
 * @apiParam (body) {Number} [exchangeRate] Exchange rate -> to calculate local purchase price
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully added watches to stock"
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse MissingInvoiceNumber
 * @apiUse MissingInvoiceDate
 * @apiUse MissingShipmentDate
 * @apiUse MissingPrice
 * @apiUse CredentialsError
 */
 module.exports.addToStock = async (req, res) => {
  const { _id: createdBy } = req.user;
  const { products, exchangeRate } = req.body;
  let { location } = req.body;

  // Check if required data has been sent
  if (!products || !products.length) throw new Error(error.MISSING_PARAMETERS);

  // Create 'toExecute' array
  const toExecute = [];

  // For each product ID sent
  for (const productId of products) {
    // Find watch in 'SoonInStock' collection
    const soonInStock = await SoonInStock.findOne({ _id: productId }).lean();

    // Check if soon in stock watch exists in DB
    if (!soonInStock) throw new Error(error.NOT_FOUND);

    // Check if soon in stock watch has 'invoiceNumber' -> required to identify the shipment it belongs to
    if (!soonInStock.invoiceNumber) throw new Error(error.MISSING_INVOICE_NUMBER);

    // Check if soon in stock watch has 'invoiceDate' -> necessary to set 'invoiceDate' for the belonging shipment
    if (!soonInStock.invoiceDate) throw new Error(error.MISSING_INVOICE_DATE);

    // Check if soon in stock watch has 'shipmentDate' -> necessary to set 'shipmentDate' for the belonging shipment
    if (!soonInStock.shipmentDate) throw new Error(error.MISSING_SHIPMENT_DATE);

    // Check if product model has 'exGenevaCHF' -> necessary to calculate KPMG reports
    if (!soonInStock.exGenevaPrice) throw new Error(error.MISSING_PRICE);

    // Find product model of the soonInStock watch that is about to be moved to stock
    const productModel = await Product.findOne({ _id: soonInStock.product }, { basicInfo: 1 }).lean();

    // Find wishlist for the soonInStock RMC
    const wishlist = await Wishlist.findOne({ rmc: soonInStock.rmc, archived: false, clients: { $elemMatch: { status: 'Active', store: soonInStock.store } } }).lean();

    // Set status
    let status = 'Stock';
    if (productModel.basicInfo.collection === 'PROFESSIONAL' || wishlist) status = 'Wishlist';
    if (soonInStock.status === 'Reserved') status = 'Reserved';
    if (soonInStock.status === 'Pre-reserved') status = 'Pre-reserved';

    // Set location
    location = location && location !== '' ? location : 'No location';

    // Set purchasePriceLocal
    const purchasePriceLocal = exchangeRate ? exchangeRate * soonInStock.exGenevaPrice : null;

    // 1. Update product -> push watch to 'boutiques.serialNumbers' array
    const updateProduct = Product.findOneAndUpdate(
      {
        'basicInfo.rmc': soonInStock.rmc,
        boutiques: { $elemMatch: { store: soonInStock.store._id } }
      },
      {
        $addToSet: {
          'boutiques.$.serialNumbers': {
            number: soonInStock.serialNumber,
            pgpReference: soonInStock.pgpReference,
            // location: 'No location',
            location,
            status,
            reservedFor: soonInStock.reservedFor,
            reservationTime: soonInStock.reservationTime,
            comment: soonInStock.comment,
            warrantyConfirmed: soonInStock.warrantyConfirmed,
            origin: 'Geneva',
            card: soonInStock.card,
            stockDate: new Date(),
            // Other features
            soldToParty: soonInStock.soldToParty,
            shipToParty: soonInStock.shipToParty,
            billToParty: soonInStock.billToParty,
            packingNumber: soonInStock.packingNumber,
            invoiceDate: soonInStock.invoiceDate,
            invoiceNumber: soonInStock.invoiceNumber,
            boxCode: soonInStock.boxCode,
            exGenevaPrice: soonInStock.exGenevaPrice,
            purchasePriceLocal,
            sectorId: soonInStock.sectorId,
            sector: soonInStock.sector,
            boitId: soonInStock.boitId,
            boitRef: soonInStock.boitRef,
            caspId: soonInStock.caspId,
            cadranDesc: soonInStock.cadranDesc,
            ldisId: soonInStock.ldisId,
            ldisq: soonInStock.ldisq,
            brspId: soonInStock.brspId,
            brspRef: soonInStock.brspRef,
            prixSuisse: soonInStock.prixSuisse,
            umIntern: soonInStock.umIntern,
            umExtern: soonInStock.umExtern,
            rfid: soonInStock.rfid,
            packingType: soonInStock.packingType
          }
        },
        $inc: { 'boutiques.$.quantity': 1 }
      },
      { new: true }
    ).lean();

    toExecute.push(updateProduct);

    // 2. Delete soon in stock product
    const deleteSoonInStock = SoonInStock.deleteOne({ _id: productId });

    toExecute.push(deleteSoonInStock);

    // Find product
    const product = await Product.findOne({ 'basicInfo.rmc': soonInStock.rmc, boutiques: { $elemMatch: { store: soonInStock.store._id } } }).lean();

    // 3. Create history log
    const newActivity = new Activity({
      type: 'Product',
      user: createdBy,
      comment: `Watch added to stock`,
      product: product._id,
      serialNumber: soonInStock.serialNumber,
    });

    toExecute.push(newActivity.save());

    // Check if shipment with soonInStock 'invoiceNumber' already exists in DB
    const shipment = await Shipment.findOne({ invoiceNumber: soonInStock.invoiceNumber, store: soonInStock.store }).lean();

    // 4. Create shipment if it does not exist in DB
    if (!shipment && soonInStock.invoiceNumber && soonInStock.invoiceDate && soonInStock.shipmentDate) {
      await new Shipment({
        type: 'arrived',
        store: soonInStock.store,
        shipmentDate: soonInStock.shipmentDate,
        invoiceDate: soonInStock.invoiceDate,
        invoiceNumber: soonInStock.invoiceNumber,
        paymentDeadline: moment(soonInStock.invoiceDate).add(75, 'days'),
        lastStockDate: new Date()
      }).save();
    }

    // 5. Update shipment
    if (soonInStock.invoiceNumber) {
      const exGenevaPrice = soonInStock.exGenevaPrice ? soonInStock.exGenevaPrice / 100 : 0;

      const updateSet = { $inc: { amountCHF: exGenevaPrice, quantity: 1 } };

      if (soonInStock.invoiceDate) {
        updateSet.invoiceDate = soonInStock.invoiceDate;
        updateSet.paymentDeadline = moment(soonInStock.invoiceDate).add(75, 'days');
        updateSet.lastStockDate = new Date();
      }

      const updateShipment = Shipment.updateOne(
        {
          store: soonInStock.store,
          invoiceNumber: soonInStock.invoiceNumber
        },
        updateSet
      );

      toExecute.push(updateShipment);
    }
  }

  // Execute
  await Promise.all(toExecute);

  return res.status(200).send({
    message: 'Successfully added watches to stock',
  });
};

/**
 * @api {get} /product/filter Get product details for filter
 * @apiVersion 1.0.0
 * @apiName getFilterDetails
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} brand Filter by brand
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned product details for filter",
   "brand": "Rolex",
   "collection": [
     "OYSTER",
     "PROFESSIONAL"
   ],
   "productLine": [
     "COSMOGRAPH DAYTONA",
     "DATEJUST PEARLMASTER",
     "DAY-DATE 36",
     "DAY-DATE 40",
     "GMT-MASTER II",
     "SUBMARINER DATE"
   ]
 }
 * @apiUse MissingParamsError
 */
module.exports.getFilterDetails = async (req, res) => {
  const { brand } = req.query;

  // Check that all requested data has been sent
  if (!brand) throw new Error(error.MISSING_PARAMETERS);

  // Find watch details for filter for specific brand
  const [collection, productLine] = await Promise.all([Product.distinct('basicInfo.collection', {
    brand,
    'basicInfo.collection': { $ne: '' }
  }), Product.distinct('basicInfo.productLine', { brand, 'basicInfo.productLine': { $ne: '' } })]);

  return res.status(200).send({
    message: 'Successfully returned product details for filter',
    brand,
    collection,
    productLine,
  });
};

/**
 * @api {get} /product Get all products
 * @apiVersion 1.0.0
 * @apiName getAllProducts
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {Number} [skip=0] Number of Products to Skip
 * @apiParam (query) {Number} [limit=50] Number of Products to Display
 * @apiParam (query) {String='Belgrade', 'Budapest', 'Porto Montenegro'} [boutique] Filter by Boutique (Store) Name
 * @apiParam (query) {String='Rolex'} [brand] Filter by Brand Type
 * @apiParam (query) {Boolean= true} [inStock] Filter only products that are In Stock (i.e. products with 'quantity' > 0)
 * @apiParam (query) {String='CELLINI', 'OYSTER', 'PROFESSIONAL'} [collection] Filter by Collection Type
 * @apiParam (query) {String} [productLine] Filter by Product Line
 * @apiParam (query) {String} [saleReference] Filter by Sale Reference
 * @apiParam (query) {String} [dial] Filter by Dial
 * @apiParam (query) {String} [bracelet] Filter by Bracelet
 * @apiParam (query) {String} [numberOfPhotos] Filter by number of photos
 * @apiParam (query) {String} [search] Search products by RMC or PGP Reference
 * @apiParam (query) {String='numberOfPhotos'} [sortBy] Sort products by specified property. NOTE: For descending order use minus sign in front of property, e.g. '-numberOfPhotos'.
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned all products",
   "count": 1,
   "results": [
     {
       "_id": "5f50e73c1e5c985976250b24",
       "status": "new",
       "brand": "Rolex",
       "boutiques": [
         {
           "quantity": 2,
           "_id": "5f5b1b6d6f4bda8815cf5484",
           "store": "5f5b1b6d6f4bda8815cf5480",
           "storeName": "Belgrade",
           "price": 1440200,
           "VATpercent": 20,
           "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
         },
         {
           "quantity": 4,
           "_id": "5f5b1b6d6f4bda8815cf5485",
           "store": "5f5b1b6d6f4bda8815cf5481",
           "storeName": "Budapest",
           "price": 1524250,
           "VATpercent": 27,
           "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
         },
         {
           "quantity": 0,
           "_id": "5f5b1b6d6f4bda8815cf5486",
           "store": "5f5b1b6d6f4bda8815cf5482",
           "storeName": "Porto Montenegro",
           "price": 1452400,
           "VATpercent": 21,
           "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
         }
       ],
       "basicInfo": {
         "rmc": "M116769TBRJ-0002",
         "collection": "PROFESSIONAL",
         "productLine": "GMT-MASTER II",
         "saleReference": "116769TBRJ",
         "materialDescription": "PAVED W-74779BRJ",
         "dial": "PAVED W",
         "bracelet": "74779BRJ",
         "box": "EN DD EMERAUDE 60",
         "exGeneveCHF": 1565800,
         "diameter": 40,
         "photos": [
           "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m116769tbrj-0002.png?impolicy=v6-upright&imwidth=420",
         ]
       },
       "wishlist": {
         "_id": "603cafe60d99a81393c7c4ec",
         "archived": true
       },
       "numberOfPhotos": 1,
       "__v": 0,
       "createdAt": "2020-09-03T12:53:16.774Z",
       "updatedAt": "2020-09-03T12:53:16.774Z"
     }
   ]
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 */
 module.exports.getAllProducts = async (req, res) => {
  let { stores: userStores } = req.user;
  const { skip = 0, boutique, brand, inStock, collection, productLine, saleReference, search, dial, bracelet, numberOfPhotos, sortBy } = req.query;
  let { limit = 50 } = req.query;

  // Check that skip and limit are sent correctly
  if (!limit || parseInt(limit, 10) > 50) limit = 50;
  if (!Number.isInteger(parseInt(skip, 10)) || !Number.isInteger(parseInt(limit, 10))) throw new Error(error.INVALID_VALUE);

  // Creat query object
  let query = {};

  // Check if 'boutique' filter has been sent
  if (boutique && !inStock) query = {};

  // Check if 'brand' filter has been sent
  if (brand) {
    if (!brandTypes.includes(brand)) throw new Error(error.INVALID_VALUE);
    query.brand = brand;
  }

  // Check if 'inStock' filter has been sent
  if (inStock && !boutique) query['boutiques.quantity'] = { $gt: 0 };
  if (inStock && boutique) query.boutiques = { $elemMatch: { quantity: { $gt: 0 }, storeName: boutique } };

  // Check if 'collection' filter has been sent
  if (collection) {
    if (!collectionTypes.includes(collection)) throw new Error(error.INVALID_VALUE);
    query['basicInfo.collection'] = collection;
  }

  // Check if 'productLine' filter has been sent
  if (productLine) query['basicInfo.productLine'] = productLine;

  // Check if 'saleReference' filter has been sent
  if (saleReference) query['basicInfo.saleReference'] = saleReference;

  // Check if 'dial' filter has been sent
  if (dial) query['basicInfo.dial'] = dial;

  // Check if 'bracelet' filter has been sent
  if (bracelet) query['basicInfo.bracelet'] = bracelet;

  // Search products by RMC or PGP Reference
  if (search) query.$or = [{ 'basicInfo.rmc': new RegExp(`.*${search}.*`, 'i') }, { 'boutiques.serialNumbers.pgpReference': new RegExp(search, 'i') }];

  // Get list of products and count
  let [listOfProducts, count] = await Promise.all([
    Product.find(query).populate('wishlist', 'archived').lean(),
    Product.countDocuments(query).lean(),
  ]);

  // Add a new numberOfPhotos field
  listOfProducts.map((p) => (p.numberOfPhotos = p.basicInfo.photos.length));

  // Filter products by numberOfPhotos
  if (numberOfPhotos) listOfProducts = listOfProducts.filter((p) => p.numberOfPhotos === parseInt(numberOfPhotos));

  // Sort products by numberOfPhotos
  if (sortBy === 'numberOfPhotos') listOfProducts = listOfProducts.sort((a, b) => a.numberOfPhotos - b.numberOfPhotos);
  if (sortBy === '-numberOfPhotos') listOfProducts = listOfProducts.sort((a, b) => b.numberOfPhotos - a.numberOfPhotos);

  userStores = userStores.map((store) => store._id.toString());

  for (let i = 0; i < listOfProducts.length; i++) {
    listOfProducts[i].boutiques = listOfProducts[i].boutiques.filter((boutique) => userStores.includes(boutique.store.toString()));
  }

  // Get 'results'
  count = listOfProducts.length;
  const results = listOfProducts.slice(Number(skip), Number(skip) + Number(limit));

  return res.status(200).send({
    message: 'Successfully returned all products',
    count,
    results,
  });
};

/**
 * @api {get} /product/:productId Get product details
 * @apiVersion 1.0.0
 * @apiName getProductDetails
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (params) {String} productId Product ID
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned product details",
   "results": {
     "_id": "5f50e73c1e5c985976250b24",
     "status": "new",
     "brand": "Rolex",
     "boutiques": [
       {
         "quantity": 2,
         "_id": "5f5b1b6d6f4bda8815cf5484",
         "store": "5f5b1b6d6f4bda8815cf5480",
         "storeName": "Belgrade",
         "price": 1440200,
         "VATpercent": 20,
         "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
       },
       {
         "quantity": 4,
         "_id": "5f5b1b6d6f4bda8815cf5485",
         "store": "5f5b1b6d6f4bda8815cf5481",
         "storeName": "Budapest",
         "price": 1524250,
         "VATpercent": 27,
         "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
       },
       {
         "quantity": 0,
         "_id": "5f5b1b6d6f4bda8815cf5486",
         "store": "5f5b1b6d6f4bda8815cf5482",
         "storeName": "Porto Montenegro",
         "price": 1452400,
         "VATpercent": 21,
         "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
       }
     ],
     "basicInfo": {
       "rmc": "M116769TBRJ-0002",
       "collection": "PROFESSIONAL",
       "productLine": "GMT-MASTER II",
       "saleReference": "116769TBRJ",
       "materialDescription": "PAVED W-74779BRJ",
       "dial": "PAVED W",
       "bracelet": "74779BRJ",
       "box": "EN DD EMERAUDE 60",
       "exGeneveCHF": 1565800,
       "diameter": 40,
       "photos": [
         "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m116769tbrj-0002.png?impolicy=v6-upright&imwidth=420",
       ]
     },
     "wishlist": [],
     "__v": 0,
     "createdAt": "2020-09-03T12:53:16.774Z",
     "updatedAt": "2020-09-03T12:53:16.774Z"
   }
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 */
module.exports.getProductDetails = async (req, res) => {
  let { stores: userStores } = req.user;
  const { productId } = req.params;

  // Find product
  const product = await Product.findOne({ _id: productId }).lean();

  // Check if product is found
  if (!product) throw new Error(error.NOT_FOUND);

  userStores = userStores.map((store) => store._id.toString());

  product.boutiques = product.boutiques.filter((boutique) => userStores.includes(boutique.store.toString()));
  product.boutiques.sort((a, b) => (a.storeName > b.storeName ? 1 : b.storeName > a.storeName ? -1 : 0));

  return res.status(200).send({
    message: 'Successfully returned product details',
    results: product,
  });
};

/**
 * @api {post} /product/:productId/photo Add product photo
 * @apiVersion 1.0.0
 * @apiName addProductPhoto
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (body) {String[]} photos Photo URL
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully added an image to the product",
   "results": {
     "_id": "5f50e73c1e5c985976250b24",
     "status": "new",
     "brand": "Rolex",
     "boutiques": [
       {
         "quantity": 2,
         "_id": "5f5b1b6d6f4bda8815cf5484",
         "store": "5f5b1b6d6f4bda8815cf5480",
         "storeName": "Belgrade",
         "price": 1440200,
         "VATpercent": 20,
         "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
       },
       {
         "quantity": 4,
         "_id": "5f5b1b6d6f4bda8815cf5485",
         "store": "5f5b1b6d6f4bda8815cf5481",
         "storeName": "Budapest",
         "price": 1524250,
         "VATpercent": 27,
         "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
       },
       {
         "quantity": 0,
         "_id": "5f5b1b6d6f4bda8815cf5486",
         "store": "5f5b1b6d6f4bda8815cf5482",
         "storeName": "Porto Montenegro",
         "price": 1452400,
         "VATpercent": 21,
         "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
       }
     ],
     "basicInfo": {
       "rmc": "M116769TBRJ-0002",
       "collection": "PROFESSIONAL",
       "productLine": "GMT-MASTER II",
       "saleReference": "116769TBRJ",
       "materialDescription": "PAVED W-74779BRJ",
       "dial": "PAVED W",
       "bracelet": "74779BRJ",
       "box": "EN DD EMERAUDE 60",
       "exGeneveCHF": 1565800,
       "diameter": 40,
       "photos": [
         "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m116769tbrj-0002.png?impolicy=v6-upright&imwidth=420",
       ]
     },
     "wishlist": [],
     "__v": 0,
     "createdAt": "2020-09-03T12:53:16.774Z",
     "updatedAt": "2020-09-03T12:53:16.774Z"
   }
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
module.exports.addProductPhoto = async (req, res) => {
  const { productId } = req.params;
  const { photos } = req.body;

  // Check if required data has been sent
  if (!photos) throw new Error(error.MISSING_PARAMETERS);

  // Update client
  const results = await Product.findOneAndUpdate({ _id: productId }, { $set: { 'basicInfo.photos': photos } }, { new: true }).lean();

  // Check if client was found
  if (!results) throw new Error(error.NOT_FOUND);

  return res.status(200).send({
    message: 'Successfully added an image to the product',
    results,
  });
};

/**
 * @api {patch} /product/:productId/photo Change Product Photos Order
 * @apiVersion 1.0.0
 * @apiName Change Product Photos Order
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (body) {String[]} photos Photo URL
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
  "message": "Successfully updated product photos"
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse InvalidValue
 * @apiUse CredentialsError
 */
module.exports.changeProductPhotos = async (req, res) => {
  const { photos } = req.body;
  const { productId } = req.params;

  // Check if required data has been sent
  if (!photos || !productId) throw new Error(error.MISSING_PARAMETERS);
  if (!isValidId(productId) || !Array.isArray(photos)) throw new Error(error.INVALID_VALUE);

  // Find product by id
  const product = await Product.findById(productId).lean();

  // Check if product was found
  if (!product) throw new Error(error.NOT_FOUND);

  // Update product with new photos
  const newProduct = await Product.updateOne(
    { _id: productId },
    { $set: { 'basicInfo.photos': photos }, },
  )

  return res.status(200).send({
    message: 'Successfully updated product photos',
  });
}

/**
 * @api {get} /product/review Get all modified products
 * @apiVersion 1.0.0
 * @apiName reviewProducts
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {Number} [skip=0] Number of Modified Products to Skip
 * @apiParam (query) {Number} [limit=50] Number of Modified Products to Display
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned all modified products",
   "count": 1,
   "results": [
     {
       "_id": "5f50e73c1e5c985976250b24",
       "status": "new",
       "brand": "Rolex",
       "boutiques": [
         {
           "quantity": 2,
           "_id": "5f5b1b6d6f4bda8815cf5484",
           "store": "5f5b1b6d6f4bda8815cf5480",
           "storeName": "Belgrade",
           "price": 1440200,
           "VATpercent": 20,
           "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
         },
         {
           "quantity": 4,
           "_id": "5f5b1b6d6f4bda8815cf5485",
           "store": "5f5b1b6d6f4bda8815cf5481",
           "storeName": "Budapest",
           "price": 1524250,
           "VATpercent": 27,
           "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
         },
         {
           "quantity": 0,
           "_id": "5f5b1b6d6f4bda8815cf5486",
           "store": "5f5b1b6d6f4bda8815cf5482",
           "storeName": "Porto Montenegro",
           "price": 1452400,
           "VATpercent": 21,
           "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
         }
       ],
       "basicInfo": {
         "rmc": "M116769TBRJ-0002",
         "collection": "PROFESSIONAL",
         "productLine": "GMT-MASTER II",
         "saleReference": "116769TBRJ",
         "materialDescription": "PAVED W-74779BRJ",
         "dial": "PAVED W",
         "bracelet": "74779BRJ",
         "box": "EN DD EMERAUDE 60",
         "exGeneveCHF": 1565800,
         "diameter": 40,
         "photos": [
           "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m116769tbrj-0002.png?impolicy=v6-upright&imwidth=420"
         ]
       },
       "wishlist": [],
       "__v": 0,
       "createdAt": "2020-09-03T12:53:16.774Z",
       "updatedAt": "2020-09-03T12:53:16.774Z"
     }
   ]
 }
 * @apiUse MissingParamsError
 */
module.exports.reviewProducts = async (req, res) => {
  const { skip = 0 } = req.query;
  let { limit = 50 } = req.query;

  // Check that skip and limit are sent correctly
  if (!limit || parseInt(limit, 10) > 50) limit = 50;
  if (!Number.isInteger(parseInt(skip, 10)) || !Number.isInteger(parseInt(limit, 10))) throw new Error(error.INVALID_VALUE);

  // Get list of modified products and count
  const [listOfModifiedProducts, count, newProducts, changedProducts, deletedProducts] = await Promise.all([
    ModifiedProduct.find().skip(parseInt(skip, 10)).limit(parseInt(limit, 10)).lean(),
    ModifiedProduct.countDocuments().lean(),
    ModifiedProduct.distinct('_id', { status: { $in: 'new' } }),
    ModifiedProduct.distinct('_id', { status: { $in: 'changed' } }),
    ModifiedProduct.distinct('_id', { status: { $in: 'deleted' } }),
  ]);

  return res.status(200).send({
    message: 'Successfully returned all modified products',
    count,
    new: newProducts.length,
    changed: changedProducts.length,
    deleted: deletedProducts.length,
    results: listOfModifiedProducts,
  });
};

/**
 * @api {get} /product/:productId/combination Get product combination list
 * @apiVersion 1.0.0
 * @apiName getProductCombinationList
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - User with permission
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (query) {String} [dial] Dial
 * @apiParam (query) {String} [bracelet] Bracelet
 * @apiParam (query) {String='Belgrade', 'Budapest', 'Porto Montenegro'} [storeName=logged_in_user's_store] Store Name (where the watch with the specified productId is stocked)
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
    "message": "Successfully returned product combination list",
    "product": {
        "_id": "5f6b476c5d79081351f78d49",
        "basicInfo": {
            "photos": [
                "http://91.226.243.34:3029/minio/download/rolex.test/elephant-1597843689196.jpg?token=",
                "http://91.226.243.34:3029/minio/download/rolex.test/rhino-1597843689197.jpg?token="
            ],
            "rmc": "M50705RBR-0004",
            "collection": "CELLINI",
            "productLine": "CLASSIC",
            "saleReference": "50705RBR",
            "materialDescription": "BLACK 11BR P-BUCKLE",
            "dial": "BLACK 11BR P",
            "bracelet": "BUCKLE",
            "box": "CELLINI L",
            "exGeneveCHF": null,
            "retailRsEUR": 22100,
            "retailHuEUR": 23300,
            "retailMneEUR": 22300,
            "diameter": 39
        },
        "status": "previous",
        "brand": "Rolex",
        "boutiques": [
            {
                "quantity": 1,
                "_id": "5f6b476c5d79081351f78d4a",
                "store": "5f3a4225ffe375404f72fb06",
                "storeName": "Belgrade",
                "serialNumbers": [
                  {
                    "number": "FR546SAG",
                    "stockDate": "2020-07-08T14:04:49.541Z"
                  }
                ]
            },
            {
                "quantity": 0,
                "_id": "5f6b476c5d79081351f78d4b",
                "store": "5f3a4225ffe375404f72fb07",
                "storeName": "Budapest",
                "serialNumbers": []
            },
            {
                "quantity": 0,
                "_id": "5f6b476c5d79081351f78d4c",
                "store": "5f3a4225ffe375404f72fb08",
                "storeName": "Porto Montenegro",
                "serialNumbers": []
            }
        ],
        "__v": 0,
        "createdAt": "2020-09-23T13:03:57.966Z",
        "updatedAt": "2020-10-06T13:52:30.203Z",
        "wishlist": "5f6b53275d79081351f84fed"
    },
    "dials": [
      "BLACK 11BR P",
      "BLACK 11BR W"
    ],
    "bracelets": [
      "BUCKLE",
      "CROCO BLACK",
      "CROCO BROWN"
    ],
    "similarProductsCount": 11,
    "similarProducts": [
      {
        "_id": "5f6b476c5d79081351f78d49",
        "basicInfo": {
          "photos": [
            "http://91.226.243.34:3029/minio/download/rolex.test/elephant-1597843689196.jpg?token=",
            "http://91.226.243.34:3029/minio/download/rolex.test/rhino-1597843689197.jpg?token="
          ],
          "rmc": "M50705RBR-0004",
          "collection": "CELLINI",
          "productLine": "CLASSIC",
          "saleReference": "50705RBR",
          "materialDescription": "BLACK 11BR P-BUCKLE",
          "dial": "BLACK 11BR P",
          "bracelet": "BUCKLE",
          "box": "CELLINI L",
          "exGeneveCHF": null,
          "retailRsEUR": 22100,
          "retailHuEUR": 23300,
          "retailMneEUR": 22300,
          "diameter": 39
        },
        "status": "previous",
        "brand": "Rolex",
        "boutiques": [
          {
            "quantity": 1,
            "_id": "5f6b476c5d79081351f78d4a",
            "store": "5f3a4225ffe375404f72fb06",
            "storeName": "Belgrade",
            "serialNumbers": [
              {
                "number": "FR546SAG",
                "stockDate": "2020-07-08T14:04:49.541Z"
              }
            ]
          },
          {
            "quantity": 0,
            "_id": "5f6b476c5d79081351f78d4b",
            "store": "5f3a4225ffe375404f72fb07",
            "storeName": "Budapest",
            "serialNumbers": []
          },
          {
            "quantity": 0,
            "_id": "5f6b476c5d79081351f78d4c",
            "store": "5f3a4225ffe375404f72fb08",
            "storeName": "Porto Montenegro",
            "serialNumbers": []
          }
        ],
        "__v": 0,
        "createdAt": "2020-09-23T13:03:57.966Z",
        "updatedAt": "2020-10-06T13:52:30.203Z",
        "wishlist": "5f6b53275d79081351f84fed"
      },
      ...
    ],
    "otherStoresResults": [
      {
        "store": "Budapest",
        "dials": [
          "PINK INDEX P"
        ],
        "bracelets": [
          "BUCKLE",
          "CROCO BLACK",
          "CROCO BROWN"
        ],
        "similarProductsCount": 6,
        "similarProducts": [
          {
            "_id": "5f6b476c5d79081351f78d79",
            "basicInfo": {
              "photos": [],
              "rmc": "M50705RBR-0007",
              "collection": "CELLINI",
              "productLine": "CLASSIC",
              "saleReference": "50705RBR",
              "materialDescription": "PINK INDEX P-BUCKLE",
              "dial": "PINK INDEX P",
              "bracelet": "BUCKLE",
              "box": "CELLINI L",
              "exGeneveCHF": null,
              "retailRsEUR": 21050,
              "retailHuEUR": 22200,
              "retailMneEUR": 21250,
              "diameter": 39
            },
            "status": "previous",
            "brand": "Rolex",
            "boutiques": [
              {
                "quantity": 0,
                "_id": "5f6b476c5d79081351f78d7a",
                "store": "5f3a4225ffe375404f72fb06",
                "storeName": "Belgrade",
                "serialNumbers": []
              },
              {
                "quantity": 1,
                "_id": "5f6b476c5d79081351f78d7b",
                "store": "5f3a4225ffe375404f72fb07",
                "storeName": "Budapest",
                "serialNumbers": [
                  {
                    "number": "FR546SAG",
                    "stockDate": "2020-07-08T14:04:49.541Z"
                  }
                ]
              },
              {
                "quantity": 0,
                "_id": "5f6b476c5d79081351f78d7c",
                "store": "5f3a4225ffe375404f72fb08",
                "storeName": "Porto Montenegro",
                "serialNumbers": []
              }
            ],
            "__v": 0,
            "createdAt": "2020-09-23T13:03:57.971Z",
            "updatedAt": "2020-10-06T13:52:30.205Z"
          },
          ...
        ]
      },
      {
        "store": "Porto Montenegro",
        "dials": [
          "RHODIUM INDEX W"
        ],
        "bracelets": [
          "BUCKLE",
          "CROCO BLACK",
          "CROCO BROWN"
        ],
        "similarProductsCount": 6,
        "similarProducts": [
          {
            "_id": "5f6b476c5d79081351f78d91",
            "basicInfo": {
              "photos": [],
              "rmc": "M50709RBR-0006",
              "collection": "CELLINI",
              "productLine": "CLASSIC",
              "saleReference": "50709RBR",
              "materialDescription": "RHODIUM INDEX W-BUCKLE",
              "dial": "RHODIUM INDEX W",
              "bracelet": "BUCKLE",
              "box": "CELLINI L",
              "exGeneveCHF": null,
              "retailRsEUR": 21050,
              "retailHuEUR": 22200,
              "retailMneEUR": 21250,
              "diameter": 39
            },
            "status": "previous",
            "brand": "Rolex",
            "boutiques": [
              {
                "quantity": 0,
                "_id": "5f6b476c5d79081351f78d92",
                "store": "5f3a4225ffe375404f72fb06",
                "storeName": "Belgrade",
                "serialNumbers": []
              },
              {
                "quantity": 0,
                "_id": "5f6b476c5d79081351f78d93",
                "store": "5f3a4225ffe375404f72fb07",
                "storeName": "Budapest",
                "serialNumbers": []
              },
              {
                "quantity": 1,
                "_id": "5f6b476c5d79081351f78d94",
                "store": "5f3a4225ffe375404f72fb08",
                "storeName": "Porto Montenegro",
                "serialNumbers": [
                  {
                    "number": "FR546SAG",
                    "stockDate": "2020-07-08T14:04:49.541Z"
                  }
                ]
              }
            ],
            "__v": 0,
            "createdAt": "2020-09-23T13:03:57.972Z",
            "updatedAt": "2020-10-06T13:52:30.222Z"
          },
          ...
        ]
      }
    ]
  }
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
module.exports.getProductCombinationList = async (req, res) => {
  const { productId } = req.params;
  const { dial, bracelet } = req.query;
  let { storeName } = req.query;

  // Find product in DB
  const product = await Product.findById(productId).lean();

  // Check if product was found
  if (!product) throw new Error(error.NOT_FOUND);

  // Set 'storeName'
  storeName = storeName ? storeName : req.user.store.name;

  // Get other stores
  const otherStores = await Store.distinct('name', { name: { $ne: storeName } });

  // Create 'query' object
  const query = {
    'basicInfo.productLine': product.basicInfo.productLine,
    'basicInfo.diameter': product.basicInfo.diameter,
    boutiques: { $elemMatch: { storeName, quantity: { $gte: 1 } } }
  };

  // Update 'query' based on the type of part that has been sent
  if (dial) query['basicInfo.dial'] = dial;
  if (bracelet) query['basicInfo.bracelet'] = bracelet;

  // Find similar watches, dials and bracelets for the specified 'productLine' and 'diameter'
  const [dials, bracelets, similarProducts, similarProductsCount] = await Promise.all([
    Product.distinct('basicInfo.dial', query),
    Product.distinct('basicInfo.bracelet', query),
    Product.find(query).lean(),
    Product.countDocuments(query).lean(),
  ]);

  // Crete 'otherStoresResults' array
  const otherStoresResults = [];

  // For each of the other stores
  for (let otherStore of otherStores) {
    // Create quety object
    const query = {
      'basicInfo.productLine': product.basicInfo.productLine,
      'basicInfo.diameter': product.basicInfo.diameter,
      boutiques: { $elemMatch: { storeName: otherStore, quantity: { $gte: 1 } } }
    };

    // Update 'query' based on the type of part that has been sent
    if (dial) query['basicInfo.dial'] = dial;
    if (bracelet) query['basicInfo.bracelet'] = bracelet;

    // Find similar watches, dials and bracelets for the specified 'productLine' and 'diameter'
    const [dials, bracelets, similarProducts, similarProductsCount] = await Promise.all([
      Product.distinct('basicInfo.dial', query),
      Product.distinct('basicInfo.bracelet', query),
      Product.find(query).lean(),
      Product.countDocuments(query).lean(),
    ]);

    // Create 'results' object
    const results = { store: otherStore, dials, bracelets, similarProductsCount, similarProducts };
    // Push 'results' object into 'otherStoresResults' array
    otherStoresResults.push(results);
  }

  return res.status(200).send({
    message: 'Successfully returned product combination list',
    product,
    dials,
    bracelets,
    similarProductsCount,
    similarProducts,
    otherStoresResults
  });
};

/**
 * @api {patch} /product/:productId/exchange Exchange products parts
 * @apiVersion 1.0.0
 * @apiName exchangeParts
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - User with permission
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (body) {String} dial Dial
 * @apiParam (body) {String} bracelet Bracelet
 * @apiParam (body) {String='Belgrade', 'Budapest', 'Porto Montenegro'} [storeName=logged_in_user's_store] Store Name (where the watch of productId is stocked)
 * @apiParam (body) {String='Belgrade', 'Budapest', 'Porto Montenegro'} [otherStoreName=storeName] Other Store Name (if the watch we are doing the exchange with is stocked in a different store than productId watch)
 * @apiParam (body) {Boolean} [save] Indicates whether to save changes into DB or not
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
    "message": "Successfully swaped product parts",
    "results": {
      "productSerialNumbers": [
        {
          "_id": "5f7f1cafed15545b8d384ef5",
          "basicInfo": {
            "photos": [
              "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m50609rbr-0010.png?impolicy=v6-upright&imwidth=420"
            ],
            "rmc": "M50609RBR-0010",
            "collection": "CELLINI",
            "productLine": "CLASSIC",
            "saleReference": "50609RBR",
            "materialDescription": "BLACK 11BR W-CROCO BROWN",
            "dial": "BLACK 11BR W",
            "bracelet": "CROCO BROWN",
            "box": "CELLINI L",
            "exGeneveCHF": null,
            "diameter": 39
          },
          "status": "new",
          "brand": "Rolex",
          "boutiques": {
            "quantity": 2,
            "_id": "5f7f1cafed15545b8d384ef6",
            "store": "5f3a4225ffe375404f72fb06",
            "storeName": "Belgrade",
            "price": 19000,
            "VATpercent": 20,
            "serialNumbers": {
              "number": "AA23BB34",
              "stockDate": "2020-02-08T14:04:49.541Z"
            }
          },
          "__v": 0,
          "createdAt": "2020-10-08T14:05:58.115Z",
          "updatedAt": "2020-10-11T20:40:22.051Z"
        },
        ...,
        {
          "_id": "5f7f1cafed15545b8d384ef5",
          "basicInfo": {
            "photos": [
              "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m50609rbr-0010.png?impolicy=v6-upright&imwidth=420"
            ],
            "rmc": "M50609RBR-0010",
            "collection": "CELLINI",
            "productLine": "CLASSIC",
            "saleReference": "50609RBR",
            "materialDescription": "BLACK 11BR W-CROCO BROWN",
            "dial": "BLACK 11BR W",
            "bracelet": "CROCO BROWN",
            "box": "CELLINI L",
            "exGeneveCHF": null,
            "diameter": 39
          },
          "status": "new",
          "brand": "Rolex",
          "boutiques": {
            "quantity": 2,
            "_id": "5f7f1cafed15545b8d384ef6",
            "store": "5f3a4225ffe375404f72fb06",
            "storeName": "Belgrade",
            "price": 19000,
            "VATpercent": 20,
            "serialNumbers": {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          },
          "__v": 0,
          "createdAt": "2020-10-08T14:05:58.115Z",
          "updatedAt": "2020-10-11T20:40:22.051Z"
        }
      ],
      "otherProductsSerialNumbers": [
        {
          "_id": "5f7f1cafed15545b8d384ebd",
          "basicInfo": {
            "photos": [
              "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m50605rbr-0014.png?impolicy=v6-upright&imwidth=420"
            ],
            "rmc": "M50605RBR-0014",
            "collection": "CELLINI",
            "productLine": "CLASSIC",
            "saleReference": "50605RBR",
            "materialDescription": "BLACK 11BR P-CROCO BLACK",
            "dial": "BLACK 11BR P",
            "bracelet": "CROCO BLACK",
            "box": "CELLINI L",
            "exGeneveCHF": null,
            "diameter": 39
          },
          "status": "new",
          "brand": "Rolex",
          "boutiques": {
            "quantity": 1,
            "_id": "5f7f1cafed15545b8d384ebe",
            "store": "5f3a4225ffe375404f72fb06",
            "storeName": "Belgrade",
            "price": 19000,
            "VATpercent": 20,
            "serialNumbers": {
              "number": "DD34FF56GG",
              "stockDate": "2019-05-08T14:05:58.112Z"
            }
          },
          "__v": 0,
          "createdAt": "2020-10-08T14:05:58.112Z",
          "updatedAt": "2020-10-11T20:40:22.051Z"
        },
        ...
      ],
      "desiredProduct": {
        "_id": "5f7f1cafed15545b8d384eb5",
        "basicInfo": {
          "photos": [
            "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m50605rbr-0013.png?impolicy=v6-upright&imwidth=420"
          ],
          "rmc": "M50605RBR-0013",
          "collection": "CELLINI",
          "productLine": "CLASSIC",
          "saleReference": "50605RBR",
          "materialDescription": "BLACK 11BR P-CROCO BROWN",
          "dial": "BLACK 11BR P",
          "bracelet": "CROCO BROWN",
          "box": "CELLINI L",
          "exGeneveCHF": null,
          "diameter": 39
        },
        "status": "new",
        "brand": "Rolex",
        "boutiques": [
          {
            "quantity": 1,
            "_id": "5f7f1cafed15545b8d384eb6",
            "store": "5f3a4225ffe375404f72fb06",
            "storeName": "Belgrade",
            "price": 19000,
            "VATpercent": 20,
            "serialNumbers": [
              {
                "_id": "5f837228b5098e4a6338dec6",
                "number": "AA23BB34",
                "stockDate": "2020-02-08T14:04:49.541Z",
                "modified": true,
                "modificationDate": "2020-10-11T20:59:20.261Z",
                "modifiedBy": "5f50e2ad8a99e565753e5f63"
              }
            ]
          },
          {
            "quantity": 0,
            "_id": "5f7f1cafed15545b8d384eb7",
            "store": "5f3a4225ffe375404f72fb07",
            "storeName": "Budapest",
            "price": 20100,
            "VATpercent": 27,
            "serialNumbers": []
          },
          {
            "quantity": 0,
            "_id": "5f7f1cafed15545b8d384eb8",
            "store": "5f3a4225ffe375404f72fb08",
            "storeName": "Porto Montenegro",
            "price": 19200,
            "VATpercent": 21,
            "serialNumbers": []
          }
        ],
        "__v": 0,
        "createdAt": "2020-10-08T14:05:58.111Z",
        "updatedAt": "2020-10-11T20:59:20.264Z"
      },
      "byProduct": {
        "_id": "5f7f1cafed15545b8d384edd",
        "basicInfo": {
          "photos": [
            "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m50609rbr-0007.png?impolicy=v6-upright&imwidth=420"
          ],
          "rmc": "M50609RBR-0007",
          "collection": "CELLINI",
          "productLine": "CLASSIC",
          "saleReference": "50609RBR",
          "materialDescription": "BLACK 11BR W-CROCO BLACK",
          "dial": "BLACK 11BR W",
          "bracelet": "CROCO BLACK",
          "box": "CELLINI L",
          "exGeneveCHF": null,
          "diameter": 39
        },
        "status": "new",
        "brand": "Rolex",
        "boutiques": [
          {
            "quantity": 2,
            "_id": "5f7f1cafed15545b8d384ede",
            "store": "5f3a4225ffe375404f72fb06",
            "storeName": "Belgrade",
            "price": 19000,
            "VATpercent": 20,
            "serialNumbers": [
              {
                "_id": "5f7f1cf9ed15545b8d3866e3",
                "number": "92A13262",
                "stockDate": "2020-07-08T14:06:49.526Z"
              },
              {
                "_id": "5f837228b5098e4a6338dec7",
                "number": "DD34FF56GG",
                "stockDate": "2019-05-08T14:05:58.112Z",
                "modified": true,
                "modificationDate": "2020-10-11T20:59:20.262Z",
                "modifiedBy": "5f50e2ad8a99e565753e5f63"
              }
            ]
          },
          {
            "quantity": 0,
            "_id": "5f7f1cafed15545b8d384edf",
            "store": "5f3a4225ffe375404f72fb07",
            "storeName": "Budapest",
            "price": 20100,
            "VATpercent": 27,
            "serialNumbers": []
          },
          {
            "quantity": 0,
            "_id": "5f7f1cafed15545b8d384ee0",
            "store": "5f3a4225ffe375404f72fb08",
            "storeName": "Porto Montenegro",
            "price": 19200,
            "VATpercent": 21,
            "serialNumbers": []
          }
        ],
        "__v": 0,
        "createdAt": "2020-10-08T14:05:58.113Z",
        "updatedAt": "2020-10-11T20:59:20.266Z"
      },
      "initialProduct": {
        "_id": "5f7f1cafed15545b8d384ef5",
        "basicInfo": {
          "photos": [
            "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m50609rbr-0010.png?impolicy=v6-upright&imwidth=420"
          ],
          "rmc": "M50609RBR-0010",
          "collection": "CELLINI",
          "productLine": "CLASSIC",
          "saleReference": "50609RBR",
          "materialDescription": "BLACK 11BR W-CROCO BROWN",
          "dial": "BLACK 11BR W",
          "bracelet": "CROCO BROWN",
          "box": "CELLINI L",
          "exGeneveCHF": null,
          "diameter": 39
        },
        "status": "new",
        "brand": "Rolex",
        "boutiques": [
          {
            "quantity": 1,
            "_id": "5f7f1cafed15545b8d384ef6",
            "store": "5f3a4225ffe375404f72fb06",
            "storeName": "Belgrade",
            "price": 19000,
            "VATpercent": 20,
            "serialNumbers": [
              {
                "number": "FR546SAG",
                "stockDate": "2020-07-08T14:04:49.541Z"
              }
            ]
          },
          {
            "quantity": 0,
            "_id": "5f7f1cafed15545b8d384ef7",
            "store": "5f3a4225ffe375404f72fb07",
            "storeName": "Budapest",
            "price": 20100,
            "VATpercent": 27,
            "serialNumbers": []
          },
          {
            "quantity": 0,
            "_id": "5f7f1cafed15545b8d384ef8",
            "store": "5f3a4225ffe375404f72fb08",
            "storeName": "Porto Montenegro",
            "price": 19200,
            "VATpercent": 21,
            "serialNumbers": []
          }
        ],
        "__v": 0,
        "createdAt": "2020-10-08T14:05:58.115Z",
        "updatedAt": "2020-10-11T20:59:20.266Z"
      },
      "exchangedProduct": {
        "_id": "5f7f1cafed15545b8d384ebd",
        "basicInfo": {
          "photos": [
            "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m50605rbr-0014.png?impolicy=v6-upright&imwidth=420"
          ],
          "rmc": "M50605RBR-0014",
          "collection": "CELLINI",
          "productLine": "CLASSIC",
          "saleReference": "50605RBR",
          "materialDescription": "BLACK 11BR P-CROCO BLACK",
          "dial": "BLACK 11BR P",
          "bracelet": "CROCO BLACK",
          "box": "CELLINI L",
          "exGeneveCHF": null,
          "diameter": 39
        },
        "status": "new",
        "brand": "Rolex",
        "boutiques": [
          {
            "quantity": 0,
            "_id": "5f7f1cafed15545b8d384ebe",
            "store": "5f3a4225ffe375404f72fb06",
            "storeName": "Belgrade",
            "price": 19000,
            "VATpercent": 20,
            "serialNumbers": []
          },
          {
            "quantity": 0,
            "_id": "5f7f1cafed15545b8d384ebf",
            "store": "5f3a4225ffe375404f72fb07",
            "storeName": "Budapest",
            "price": 20100,
            "VATpercent": 27,
            "serialNumbers": []
          },
          {
            "quantity": 0,
            "_id": "5f7f1cafed15545b8d384ec0",
            "store": "5f3a4225ffe375404f72fb08",
            "storeName": "Porto Montenegro",
            "price": 19200,
            "VATpercent": 21,
            "serialNumbers": []
          }
        ],
        "__v": 0,
        "createdAt": "2020-10-08T14:05:58.112Z",
        "updatedAt": "2020-10-11T20:59:20.266Z"
      }
    }
  }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
module.exports.exchangeParts = async (req, res) => {
  const { productId } = req.params;
  const { dial, bracelet, save } = req.body;
  let { storeName, otherStoreName } = req.body;
  const { _id: userId } = req.user;

  // Check if all required parameters have been sent
  if (!dial || !bracelet) throw new Error(error.MISSING_PARAMETERS);

  // Set 'storeName'
  storeName = storeName ? storeName : req.user.store.name;

  // Find the initial product's serial numbers (watches) in DB
  const productSerialNumbers = await Product.aggregate([
    {
      $match: {
        _id: ObjectId(productId),
        boutiques: { $elemMatch: { storeName, quantity: { $gte: 1 } } }
      }
    },
    { $unwind: '$boutiques' },
    { $match: { 'boutiques.storeName': storeName } },
    { $unwind: '$boutiques.serialNumbers' },
    { $sort: { 'boutiques.serialNumbers.stockDate': 1 } }
  ]);

  // Check if (watches) serial numbers were found
  if (productSerialNumbers.length === 0) throw new Error(error.NOT_FOUND);

  // Take the (watch) serial number that was stocked the earliest -> that one will be modified
  const [serialNumber] = productSerialNumbers;

  // Set the 'otherStoreName'
  otherStoreName = otherStoreName ? otherStoreName : storeName;

  // Create 'query' object -> to find the other exchanging watch
  const query = {
    'basicInfo.productLine': serialNumber.basicInfo.productLine,
    'basicInfo.diameter': serialNumber.basicInfo.diameter,
    boutiques: { $elemMatch: { storeName: otherStoreName, quantity: { $gte: 1 } } }
  };

  // Update 'query' based on the type of part that has been sent
  if (serialNumber.basicInfo.dial !== dial) {
    query['basicInfo.dial'] = dial;
  } else if (serialNumber.basicInfo.bracelet !== bracelet) {
    query['basicInfo.bracelet'] = bracelet;
  }

  // Find compatible products and their serial numbers
  const otherProductsSerialNumbers = await Product.aggregate([
    { $match: query },
    { $unwind: '$boutiques' },
    { $match: { 'boutiques.storeName': otherStoreName } },
    { $unwind: '$boutiques.serialNumbers' },
    { $sort: { 'boutiques.serialNumbers.stockDate': 1 } }
  ]);

  // Check if compatible (watches) serial numbers were found
  if (otherProductsSerialNumbers.length === 0) throw new Error(error.NOT_FOUND);

  // Among all compatible (watches) serial numbers take the one that was stocked the earliest -> that one will be modified
  const [otherSerialNumber] = otherProductsSerialNumbers;

  // Create 'query2' object -> to find product RMC of the newly created (desired) watch
  const query2 = {
    'basicInfo.productLine': serialNumber.basicInfo.productLine,
    'basicInfo.diameter': serialNumber.basicInfo.diameter,
    boutiques: { $elemMatch: { storeName } },
    'basicInfo.dial': dial,
    'basicInfo.bracelet': bracelet
  };

  // Set 'query3' object -> to find product RMC of newly created (byProduct) watch
  const query3 = { ...query2 };
  query3.boutiques = { $elemMatch: { storeName: otherStoreName } };

  // Update 'query3' based on the type of part that will be changed between two selected watches ('serialNumber' and 'otherSerialNumber')
  if (serialNumber.basicInfo.dial !== dial) {
    query3['basicInfo.dial'] = serialNumber.basicInfo.dial;
    query3['basicInfo.bracelet'] = otherSerialNumber.basicInfo.bracelet;
  } else if (serialNumber.basicInfo.bracelet !== bracelet) {
    query3['basicInfo.dial'] = otherSerialNumber.basicInfo.dial;
    query3['basicInfo.bracelet'] = serialNumber.basicInfo.bracelet;
  }

  // Create update sets
  let addToSet1 = {};
  let addToSet2 = {};
  let removeFromSet1 = {};
  let removeFromSet2 = {};

  // Check if 'save' has been sent as true
  if (save) {
    addToSet1 = {
      $addToSet:
        {
          'boutiques.$.serialNumbers': {
            number: serialNumber.boutiques.serialNumbers.number,
            stockDate: serialNumber.boutiques.serialNumbers.stockDate,
            status: serialNumber.boutiques.serialNumbers.status,
            location: serialNumber.boutiques.serialNumbers.location,
            origin: serialNumber.boutiques.serialNumbers.origin,
            modified: true,
            modificationDate: new Date(),
            modifiedBy: userId
          }
        },
      $inc: { 'boutiques.$.quantity': 1 }
    };
    addToSet2 = {
      $addToSet:
        {
          'boutiques.$.serialNumbers': {
            number: otherSerialNumber.boutiques.serialNumbers.number,
            stockDate: otherSerialNumber.boutiques.serialNumbers.stockDate,
            status: otherSerialNumber.boutiques.serialNumbers.status,
            location: otherSerialNumber.boutiques.serialNumbers.location,
            origin: otherSerialNumber.boutiques.serialNumbers.origin,
            modified: true,
            modificationDate: new Date(),
            modifiedBy: userId
          }
        },
      $inc: { 'boutiques.$.quantity': 1 }
    };
    removeFromSet1 = {
      $pull: { 'boutiques.$.serialNumbers': { number: serialNumber.boutiques.serialNumbers.number } },
      $inc: { 'boutiques.$.quantity': -1 }
    };
    removeFromSet2 = {
      $pull: { 'boutiques.$.serialNumbers': { number: otherSerialNumber.boutiques.serialNumbers.number } },
      $inc: { 'boutiques.$.quantity': -1 }
    };
  }

  // Find RMCs of newly created watches and push 'serialNumber' and 'otherSerialNumber' into them, also pull 'serialNumber' and 'otherSerialNUmber' from their original RMCs
  const [rmc3, rmc4, rmc1, rmc2] = await Promise.all([
    // Push modified 'serialNumber' to belonging RMC -> this is the desired watch (first result of modification)
    Product.findOneAndUpdate(
      query2,
      addToSet1,
      { new: true }
    ).lean(),
    // Push modified 'otherSerialNumber' watch to belonging RMC -> this is the by-product watch, i.e. not the desired one (second result of modification)
    Product.findOneAndUpdate(
      query3,
      addToSet2,
      { new: true }
    ).lean(),
    // Pull 'serialNumber' from its original (initial) product RMC
    Product.findOneAndUpdate(
      { _id: ObjectId(productId), boutiques: { $elemMatch: { storeName } } },
      removeFromSet1,
      { new: true }
    ).lean(),
    // Pull 'otherSerialNumber' from its original RMC
    Product.findOneAndUpdate(
      { _id: ObjectId(otherSerialNumber._id), boutiques: { $elemMatch: { storeName: otherStoreName } } },
      removeFromSet2,
      { new: true }
    ).lean(),
  ]);

  return res.status(200).send({
    message: 'Successfully swaped product parts',
    results: {
      productSerialNumbers,
      otherProductsSerialNumbers,
      desiredProduct: rmc3,
      byProduct: rmc4,
      initialProduct: rmc1,
      exchangedProduct: rmc2
    }
  });
};

/**
 * @api {get} /product/:productId/stock Get all product watches
 * @apiVersion 1.0.0
 * @apiName getProductWatches
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - User with permission
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (query) {String='Belgrade', 'Budapest', 'Porto Montenegro', 'Global'} [storeName=logged_in_user's_store] Store Name (where the watch is stocked)
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
    "message": "Successfully returned list of watches",
    "watchesCount": 2,
    "soonInStockWatchesCount": 2,
    "count": 4,
    "results": [
      {
        "_id": "5f9aef2d923676df0f979ba7",
        "basicInfo": {
          "photos": [
            "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m114200-0020.png?impolicy=v6-upright&imwidth=420"
          ],
          "rmc": "M114200-0020",
          "collection": "OYSTER",
          "productLine": "PERPETUAL",
          "saleReference": "114200",
          "materialDescription": "RED GRAPE INDEX W-70190",
          "dial": "RED GRAPE INDEX W",
          "bracelet": "70190",
          "box": "OYSTER S",
          "exGeneveCHF": null,
          "diameter": 34
        },
        "status": "previous",
        "brand": "Rolex",
        "boutiques": {
          "quantity": 1,
          "_id": "5f9aef2d923676df0f979baa",
          "store": "5f3a4225ffe375404f72fb08",
          "storeName": "Porto Montenegro",
          "price": 4700,
          "VATpercent": 21,
          "serialNumbers": {
            "modified": false,
            "_id": "5f9ee89552944f14e57301bb",
            "number": "TIVAT-301",
            "stockDate": "2020-11-01T16:55:49.708Z",
            "origin": "Geneva",
            "location": "",
            "status": "Stock"
          }
        },
        "__v": 0,
        "createdAt": "2020-10-29T16:36:26.338Z",
        "updatedAt": "2020-11-01T16:55:49.709Z"
      },
      {
        "_id": "5f9ee7f152944f14e57301b3",
        "rmc": "M114200-0020",
        "product": "5f9aef2d923676df0f979ba7",
        "serialNumber": "TIVAT-300",
        "dial": "RED GRAPE INDEX W",
        "store": "5f3a4225ffe375404f72fb08",
        "origin": "Geneva",
        "location": "To be transferred",
        "status": "Standby"
        "__v": 0,
        "createdAt": "2020-11-01T16:53:05.985Z",
        "updatedAt": "2020-11-01T16:53:05.985Z"
      },
      ...
    ]
  }
 * @apiUse InvalidValue
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
module.exports.getProductWatches = async (req, res) => {
  const { productId } = req.params;
  let { skip = 0, limit = 50, storeName } = req.query;

  // Validate 'skip' and 'limit'
  if (!Number.isInteger(parseInt(skip, 10)) || !Number.isInteger(parseInt(limit, 10))) throw new Error(error.INVALID_VALUE);

  // Create 'storeId' string
  let storeId = '';

  // Check if 'storeName' exists in DB
  if (storeName && storeName !== 'Global') {
    const store = await Store.findOne({ name: storeName }).lean();
    if (!store) throw new Error(error.NOT_FOUND);
    storeId = store._id;
  }

  // Check if 'storeName' has been sent
  if (!storeName) storeId = req.user.store._id;

  // Set 'match' and 'query' objects
  let match = { 'boutiques.store': storeId };
  let query = { product: productId, store: storeId };

  // Check if 'storeName' has been sent as 'Global'
  if (storeName === 'Global') {
    match = { _id: ObjectId(productId) };
    query = { product: productId };
  }

  // Find the initial product's serial numbers (watches) in DB
  const [watches, soonInStockWatches] = await Promise.all([
    Product.aggregate([
      { $match: { _id: ObjectId(productId) } },
      { $unwind: '$boutiques' },
      { $match: match },
      { $unwind: '$boutiques.serialNumbers' },
      { $sort: { 'boutiques.serialNumbers.stockDate': 1 } }
    ]),
    SoonInStock.find(query).lean(),
  ]);

  // Group together all watches
  const allWatches = [...watches, ...soonInStockWatches];

  // Calculate 'count'
  const count = allWatches.length;
  const watchesCount = watches.length;
  const soonInStockWatchesCount = soonInStockWatches.length;

  // Get 'results'
  const results = allWatches.slice(Number(skip), Number(skip) + Number(limit));

  return res.status(200).send({
    message: 'Successfully returned list of watches',
    watchesCount,
    soonInStockWatchesCount,
    count,
    results
  });
};

/**
 * @api {get} /product/:productId/serialnumber Get watch details
 * @apiVersion 1.0.0
 * @apiName getWatchDetails
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - User with permission
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (query) {String} serialNumber Serial Number
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
    "message": "Successfully returned watch details",
    "results": [
      {
        "_id": "5f9aef2d923676df0f979ba7",
        "basicInfo": {
          "photos": [
            "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m114200-0020.png?impolicy=v6-upright&imwidth=420"
          ],
          "rmc": "M114200-0020",
          "collection": "OYSTER",
          "productLine": "PERPETUAL",
          "saleReference": "114200",
          "materialDescription": "RED GRAPE INDEX W-70190",
          "dial": "RED GRAPE INDEX W",
          "bracelet": "70190",
          "box": "OYSTER S",
          "exGeneveCHF": null,
          "diameter": 34
        },
        "boutiques": {
          "quantity": 1,
          "_id": "5f9aef2d923676df0f979ba8",
          "store": "5f3a4225ffe375404f72fb06",
          "storeName": "Belgrade",
          "price": 4600,
          "VATpercent": 20,
          "serialNumbers": {
            "modified": false,
            "_id": "5f9af8c7d32dace104830349",
            "number": "BEOGRAD-001",
            "stockDate": "2020-10-29T17:15:51.610Z",
            "origin": "Geneva",
            "status": "Reserved",
            "location": "Rear window",
            "reservedFor": {
              "_id": "5f9aa99e504b82b98becd099",
              "fullName": "Sloba Prvi Nikolić",
              "photo": "https://pgp-rolex.com/images/photo1.jpeg"
            },
            "reservationTime": "2020-11-11T00:00:00.000Z"
          }
        },
      },
      [
        {
          "_id": "5f9afb74d4c826e3505040f8",
          "type": "Product",
          "user": {
            "_id": "5f50e2ad8a99e565753e5f5a",
            "name": "user1"
          },
          "client": null,
          "comment": "Changed status from 'Stock' to 'Paid'.",
          "wishlist": null,
          "product": "5f9aef2d923676df0f979ba7",
          "serialNumber": "BEOGRAD-001",
          "createdAt": "2020-10-29T17:27:16.851Z",
          "updatedAt": "2020-10-29T17:27:16.851Z",
          "__v": 0
        },
        ...
      ]
    ],
    "productBoutiques": [
      {
          "storeName": "Belgrade",
          "price": 7450,
          "priceLocal": 969000
      },
      {
          "storeName": "Budapest",
          "price": 7850,
          "priceLocal": 2944000
      },
      {
          "storeName": "Porto Montenegro",
          "price": 7500,
          "priceLocal": 7500
      }
    ]
  }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
module.exports.getWatchDetails = async (req, res) => {
  const { productId } = req.params;
  const { serialNumber } = req.query;

  // Check if required parameter has been sent
  if (!serialNumber) throw new Error(error.MISSING_PARAMETERS);

  // Find product that contains sent 'serialNumber' and all activities related to the sent 'serialNumber'
  const [products, activities, productBoutiques] = await Promise.all([
    Product.aggregate([
      { $match: { _id: ObjectId(productId) } },
      { $unwind: '$boutiques' },
      { $unwind: '$boutiques.serialNumbers' },
      { $match: { 'boutiques.serialNumbers.number': serialNumber } },
    ]),
    Activity.find({
      type: 'Product',
      product: productId,
      serialNumber
    })
      .populate('user', 'name')
      .sort('-createdAt')
      .lean(),
    Product.findOne({ _id: productId }, {
      'boutiques.price': 1,
      'boutiques.priceLocal': 1,
      'boutiques.storeName': 1
    }).lean()
  ]);

  const [product] = products;

  // Check if product was found
  if (!product) throw new Error(error.NOT_FOUND);

  // Check if 'reservedFor' exists on fetched 'serialNumber'
  let client = {};
  const clientId = product.boutiques.serialNumbers.reservedFor;
  if (clientId) {
    client = await Client.findOne({ _id: clientId }, { 'fullName': 1, 'photo': 1 }).lean();
  }

  product.boutiques.serialNumbers.reservedFor = client;

  return res.status(200).send({
    message: 'Successfully returned watch details',
    results: [product, activities],
    productBoutiques: productBoutiques.boutiques
  });
};

/**
 * @api {patch} /product/:productId/serialnumber/change-store Change watch store
 * @apiVersion 1.0.0
 * @apiName changeStore
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - User with permission
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (body) {String} serialNumber Serial Number
 * @apiParam (body) {String='Belgrade', 'Budapest', 'Porto Montenegro'} newStore Store Name of new watch location
 * @apiParam (body) {String} [comment] Comment
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
    "message": "Successfully changed watch store",
    "results": {
      "_id": "5f88a905c17ab3217e654c9a",
      "basicInfo": {
        "photos": [
          "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m50605rbr-0003.png?impolicy=v6-upright&imwidth=420"
        ],
        "rmc": "M50605RBR-0003",
        "collection": "CELLINI",
        "productLine": "CLASSIC",
        "saleReference": "50605RBR",
        "materialDescription": "BLACK 11BR P-BUCKLE",
        "dial": "BLACK 11BR P",
        "bracelet": "BUCKLE",
        "box": "CELLINI L",
        "exGeneveCHF": null,
        "diameter": 39
      },
      "status": "previous",
      "brand": "Rolex",
      "boutiques": [
        {
          "quantity": 0,
          "_id": "5f88a905c17ab3217e654c9b",
          "store": "5f3a4225ffe375404f72fb06",
          "storeName": "Belgrade",
          "price": 19000,
          "VATpercent": 20,
          "serialNumbers": []
        },
        {
          "quantity": 1,
          "_id": "5f88a905c17ab3217e654c9c",
          "store": "5f3a4225ffe375404f72fb07",
          "storeName": "Budapest",
          "price": 20100,
          "VATpercent": 27,
          "serialNumbers": [
            {
              "modified": false,
              "logs": [
                "5f921edd2bfff13c916bcd2e",
                "5f921fe07bd4e73d0eabfb53",
                "5f92206eb0bb663d6ed3fdfd",
                "5f9220b25fb65e3dab38b309",
                "5f9223d361a30c405b64c1cd"
              ],
              "_id": "5f9223d361a30c405b64c1ce",
              "number": "BEOGRAD279",
              "stockDate": "2020-10-15T19:59:44.780Z",
              "origin": "Porto Montenegro"
            }
          ]
        },
        {
          "quantity": 0,
          "_id": "5f88a905c17ab3217e654c9d",
          "store": "5f3a4225ffe375404f72fb08",
          "storeName": "Porto Montenegro",
          "price": 19200,
          "VATpercent": 21,
          "serialNumbers": []
        }
      ],
      "__v": 0,
      "createdAt": "2020-10-15T19:55:03.230Z",
      "updatedAt": "2020-10-23T00:29:07.754Z"
    }
  }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse InvalidValue
 * @apiUse NotAcceptable
 * @apiUse CredentialsError
 */
module.exports.changeStore = async (req, res) => {
  const { productId } = req.params;
  const { serialNumber, newStore } = req.body;
  let { comment: sentComment } = req.body;
  const { _id: userId } = req.user;

  // Check if all required parameters have been sent
  if (!serialNumber || !newStore) throw new Error(error.MISSING_PARAMETERS);

  // Find product that contains sent 'serialNumber'
  const [product, stores] = await Promise.all([
    Product.findOne({ _id: productId, 'boutiques.serialNumbers.number': serialNumber }).lean(),
    Store.find({}).lean()
  ]);

  // Check if product was found
  if (!product) throw new Error(error.NOT_FOUND);

  // Get store names
  const storeNames = [];
  for (let store of stores) storeNames.push(store.name);

  // Check if 'newStore' is valid store name
  if (!storeNames.includes(newStore)) throw new Error(error.INVALID_VALUE);

  // Get related serial number object in product 'boutiques' array
  let watch = {};
  let currentStore = '';

  for (let boutique of product.boutiques) {
    const watchObjects = boutique.serialNumbers.filter(serNumber => serNumber.number === serialNumber);
    if (watchObjects.length > 0) {
      [watch] = watchObjects;
      currentStore = boutique.storeName;
    }
  }

  // Check if attempting to move watch to the store it's already stocked in (this causes a bug, i.e. removes the watch from that store)
  if (currentStore === newStore) throw new Error(error.NOT_ACCEPTABLE);

  // Set new 'origin'
  watch.origin = currentStore;

  // Create new activity
  let manuallyAdded = true;

  if (!sentComment) {
    sentComment = '';
    manuallyAdded = false;
  }

  const comment = `Changed store from '${currentStore}' to '${newStore}'. ${sentComment}`;
  const newActivity = createActivity('Product', userId, null, comment, null, productId, serialNumber, null, new Date(), manuallyAdded);

  // Update product -> push serial number object to new location
  await Promise.all([
    Product.updateOne(
      { _id: productId, boutiques: { $elemMatch: { storeName: newStore } } },
      {
        $addToSet:
          { 'boutiques.$.serialNumbers': watch },
        $inc: { 'boutiques.$.quantity': 1 },
      },
    ),
    newActivity.save()
  ])

  // Update product -> remove serial number object from current location
  const results = await Product.findOneAndUpdate(
    { _id: productId, boutiques: { $elemMatch: { storeName: currentStore } } },
    {
      $pull: { 'boutiques.$.serialNumbers': { number: serialNumber } },
      $inc: { 'boutiques.$.quantity': -1 }
    },
    { new: true }
  ).lean();

  return res.status(200).send({
    message: 'Successfully changed watch store',
    results
  });
};

/**
 * @api {patch} /product/:productId/serialnumber Edit watch
 * @apiVersion 1.0.0
 * @apiName editWatch
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - User with permission
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (body) {String} serialNumber Serial Number
 * @apiParam (body) {String} [location] Internal location of a watch
 * @apiParam (body) {String='Stock', 'Standby', 'Consignment', 'Display only', 'Reserved', 'Wishlist', 'In transit', 'Paid', 'Pre-reserved'} [status] Watch status
 * @apiParam (body) {String} [reservedFor] Client ID (watch is reserved for)
 * @apiParam (body) {date} [reservationTime] Date and time when reservation of a watch expires
 * @apiParam (body) {String} [comment] Comment
 * @apiParam (body) {String} [adjustedSize] Adjusted size (for rings)
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
    "message": "Successfully updated watch",
    "results": {
      "_id": "5f88a905c17ab3217e654ca2",
      "basicInfo": {
        "photos": [
          "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m50605rbr-0008.png?impolicy=v6-upright&imwidth=420"
        ],
        "rmc": "M50605RBR-0008",
        "collection": "CELLINI",
        "productLine": "CLASSIC",
        "saleReference": "50605RBR",
        "materialDescription": "PINK INDEX P-BUCKLE",
        "dial": "PINK INDEX P",
        "bracelet": "BUCKLE",
        "box": "CELLINI L",
        "exGeneveCHF": null,
        "diameter": 39
      },
      "status": "previous",
      "brand": "Rolex",
      "boutiques": [
        {
          "quantity": 0,
          "_id": "5f88a905c17ab3217e654ca3",
          "store": "5f3a4225ffe375404f72fb06",
          "storeName": "Belgrade",
          "price": 17950,
          "VATpercent": 20,
          "serialNumbers": []
        },
        {
          "quantity": 1,
          "_id": "5f88a905c17ab3217e654ca4",
          "store": "5f3a4225ffe375404f72fb07",
          "storeName": "Budapest",
          "price": 19000,
          "VATpercent": 27,
          "serialNumbers": [
            {
              "modified": false,
              "reservedFor": {
                "_id": "5f3d2c54448c653d018e790d",
                "fullName": "PSP Farman AD"
              },
              "_id": "5f88aa9e86506424744c0a82",
              "reservationTime": "2020-12-11T23:00:00.000Z",
              "number": "BUDAPEST279",
              "stockDate": "2020-10-15T20:01:34.337Z",
              "status": "Reserved",
              "origin": "Geneva",
              "location": ""
            }
          ]
        },
        {
            "quantity": 1,
            "_id": "5f88a905c17ab3217e654ca5",
            "store": "5f3a4225ffe375404f72fb08",
            "storeName": "Porto Montenegro",
            "price": 18150,
            "VATpercent": 21,
            "serialNumbers": [
              {
                "modified": true,
                "_id": "5f88b28b9a0cc528f0b24213",
                "number": "TIVAT279",
                "stockDate": "2020-10-15T20:01:59.764Z",
                "modificationDate": "2020-10-15T20:35:23.537Z",
                "modifiedBy": "5f50e2ad8a99e565753e5f63",
                "origin": "Geneva"
              }
            ]
        }
      ],
      "__v": 0,
      "createdAt": "2020-10-15T19:55:03.231Z",
      "updatedAt": "2020-10-27T13:22:10.847Z"
    }
  }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse InvalidValue
 * @apiUse CredentialsError
 */
module.exports.editWatch = async (req, res) => {
  const { productId } = req.params;
  const { serialNumber, location, status, comment, reservedFor, reservationTime, adjustedSize } = req.body;
  const { _id: userId } = req.user;

  // Check if all required parameters have been sent
  if (!serialNumber || (!location && !status && !reservedFor && !reservationTime && !comment)) throw new Error(error.MISSING_PARAMETERS);

  // Find product that contains sent 'serialNumber'
  const product = await Product.findOne({
    _id: productId,
    'boutiques.serialNumbers.number': serialNumber
  }).populate('boutiques.serialNumbers.reservedFor').lean();

  // Check if product was found
  if (!product) throw new Error(error.NOT_FOUND);

  // Check if client exists
  if (reservedFor) {
    const client = await Client.findById(reservedFor).lean();
    if (!client) throw new Error(error.NOT_FOUND);
  }

  // Validate watch status
  if (status && !statuses.includes(status)) throw new Error(error.INVALID_VALUE);

  // Find related serial number (watch) object in product's 'boutiques' array
  let watch = {};
  let boutiqueName = '';

  for (let boutique of product.boutiques) {
    const watchObjects = boutique.serialNumbers.filter(serNumber => serNumber.number === serialNumber);
    if (watchObjects.length > 0) {
      [watch] = watchObjects;
      boutiqueName = boutique.storeName;
    }
  }

  // Create 'updateWatch' object
  let updatedWatch = { ...watch };

  // Update 'updatedWatch' object
  let confirmedReservation = false;
  if (location) updatedWatch.location = location;
  if (location === '') updatedWatch.location = '';
  if (comment) updatedWatch.comment = comment;
  if (comment === '') updatedWatch.comment = '';
  if (adjustedSize) updatedWatch.adjustedSize = adjustedSize;
  if (adjustedSize === '') updatedWatch.adjustedSize = '';
  if (reservedFor || reservedFor === null) updatedWatch.reservedFor = reservedFor;
  if (reservationTime || reservationTime === '') updatedWatch.reservationTime = reservationTime;
  if (status) {
    updatedWatch.status = status;
    if (status === 'Reserved' && watch.status === 'Pre-reserved') {
      confirmedReservation = true;
    } else if (status !== watch.status) {
      updatedWatch.previousStatus = watch.status;
    }
    if (status !== 'Reserved' && status !== 'Pre-reserved') {
      updatedWatch.reservedFor = null;
      updatedWatch.reservationTime = null;
    }
  }

  // Update product -> i.e. update watch inside product
  const results = await Product.findOneAndUpdate(
    { _id: productId },
    { $set: { 'boutiques.$[].serialNumbers.$[j]': updatedWatch } },
    { arrayFilters: [{ 'j.number': serialNumber }], new: true }
  )
    .populate('boutiques.serialNumbers.reservedFor', 'fullName')
    .lean();

  // Create 'changedFields' array
  const changedFields = [];

  // Get related serial number (watch) object in results 'boutiques' array
  let resultsWatch = {};

  for (let boutique of results.boutiques) {
    const watchObjects = boutique.serialNumbers.filter(serNumberObj => serNumberObj.number === serialNumber);
    if (watchObjects.length > 0) {
      [resultsWatch] = watchObjects;
    }
  }

  // Create 'toExecute' array
  const toExecute = [];

  // Detect changes in 'location', 'status' and 'comment' -> in order to create history logs
  if (resultsWatch.location !== watch.location) changedFields.push(`Changed location from '${watch.location}' to '${resultsWatch.location}'.`);
  if (resultsWatch.status !== watch.status) changedFields.push(`Changed status from '${watch.status}' to '${resultsWatch.status}'.`);
  if (resultsWatch.comment !== watch.comment) resultsWatch.comment ? changedFields.push(`New comment: '${resultsWatch.comment}'.`) : changedFields.push(`Comment removed.`);
  if (resultsWatch.adjustedSize !== watch.adjustedSize) watch.adjustedSize ? changedFields.push(`Changed adjusted ring size from '${watch.adjustedSize}' to '${resultsWatch.adjustedSize}'.`) : changedFields.push(`Adjusted ring size: ${resultsWatch.adjustedSize}`);


  // Detect 'reservedFor' changes -> create touchpoints for clients journeys
  if (resultsWatch.reservedFor && (!watch.reservedFor)) {
    changedFields.push(`Made ${resultsWatch.status === 'Pre-reserved' ? 'pre-reservation' : 'reservation'} for '${resultsWatch.reservedFor.fullName}'.`);
  }
  if (!resultsWatch.reservedFor && watch.reservedFor) {
    changedFields.push(`Cancelled ${resultsWatch.previousStatus === 'Pre-reserved' ? 'pre-reservation' : 'reservation'} for '${watch.reservedFor.fullName}'.`);
    if (resultsWatch.reservedFor && watch.reservedFor && (resultsWatch.reservedFor._id.toString() !== watch.reservedFor._id.toString())) {
      changedFields.push(`Changed ${resultsWatch.status === 'Pre-reserved' ? 'pre-reservation' : 'reservation'} from '${watch.reservedFor.fullName}' to '${resultsWatch.reservedFor.fullName}'.`);
    }
  }
  if (resultsWatch.reservedFor && watch.reservedFor && confirmedReservation) {
    changedFields.push(`Confirmed pre-reservation for '${resultsWatch.reservedFor.fullName}'.`);
  }

  // Detect 'reservationTime' changes -> in order to create history logs
  if (resultsWatch.reservationTime && !watch.reservationTime) changedFields.push(`Reservation time set to '${moment(resultsWatch.reservationTime, 'YYYY-MM-DD').format('DD/MM/YYYY')}'.`);
  if (!resultsWatch.reservationTime && watch.reservationTime) changedFields.push(`Reservation time previously set to '${moment(watch.reservationTime).format('DD/MM/YYYY')}' has been removed.`);
  if (resultsWatch.reservationTime && watch.reservationTime && new Date(resultsWatch.reservationTime).toString() !== watch.reservationTime.toString()) changedFields.push(`Changed reservation time from '${moment(watch.reservationTime).format('DD/MM/YYYY')}' to '${moment(reservationTime, 'YYYY-MM-DD').format('DD/MM/YYYY')}'.`);

  // If there were any changes include those in the log comment
  const logComment = (changedFields.length > 0) ? `${changedFields.join(' ')}` : '';

  // Create new activity -> for watch log history
  if (changedFields.length > 0) {
    const newActivity = createActivity('Product', userId, null, logComment, null, productId, serialNumber);
    toExecute.push(newActivity.save());
  }

  // Execute
  await Promise.all(toExecute);

  return res.status(200).send({
    message: 'Successfully updated watch',
    results
  });
};

/**
 * @api {patch} /product/:productId Edit product
 * @apiVersion 1.0.0
 * @apiName editProduct
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - User with permission
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (body) {String='Messika', 'Roberto Coin', 'Petrovic Diamonds'} brand Brand
 * @apiParam (body) {String} [rmc] RMC
 * @apiParam (body) {String} [jewelryType] Jewelry type
 * @apiParam (body) {String} [materialDescription] Material description
 * @apiParam (body) {String[]} [materials] Materials array
 * @apiParam (body) {String} [size] Size
 * @apiParam (body) {String} [weight] Weight
 * @apiParam (body) {Number} [stonesQty] Number of stones
 * @apiParam (body) {Number} [allStonesWeight] Total weight of stones
 * @apiParam (body) {String} [brilliants] Brilliants
 * @apiParam (body) {String} [diaGia] Diamond quality
 * @apiParam (body) {Object[]} [stones] Stones array
 * @apiParam (body) {String} [stone.type] Stone type
 * @apiParam (body) {Number} [stone.quantity] Stone type quantity
 * @apiParam (body) {String} [stone.stoneTypeWeight] Stone type weight
 * @apiParam (body) {Object[]} [diamonds] Diamonds array
 * @apiParam (body) {Number} [diamond.quantity] Diamond quantity
 * @apiParam (body) {String} [diamond.carat] Diamond carat
 * @apiParam (body) {String} [diamond.color] Diamond color
 * @apiParam (body) {String} [diamond.clarity] Diamond clarity
 * @apiParam (body) {String} [diamond.shape] Diamond shape
 * @apiParam (body) {String} [diamond.cut] Diamond cut
 * @apiParam (body) {String} [diamond.polish] Diamond polish
 * @apiParam (body) {String} [diamond.symmetry] Diamond symmetry
 * @apiParam (body) {String[]} [diamond.giaReports[]] Diamond giaReports array
 * @apiParam (body) {String[]} [diamond.giaReportsUrls[]] Diamond giaReportsUrls array
 * @apiParam (body) {Number} [purchasePrice] Purchase price
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
  {
    "message": "Successfully updated product",
    "results": {
      "_id": "60aba936d80faa257a200eb6",
      "basicInfo": {
        "photos": [
          "ADV777CL3007RB-1622905924159.jpeg",
          "ADV777CL3007 copy-LR-1200-1625692248659.jpg"
        ],
        "materials": [
          "Blue gold",
          "White gold"
        ],
        "rmc": "123456",
        "collection": "Palazzo Ducale",
        "saleReference": "ADV777CL3007",
        "jewelryType": "collar",
        "purchasePrice": 100,
        "weight": "1.234",
        "stonesQty": 4,
        "allStonesWeight": 2.345,
        "stones": [
          {
            "_id": "61813e7282a9134e32dbdc96",
            "type": "Ruby",
            "quantity": 2,
            "stoneTypeWeight": "1.58"
          }
        ],
        "diaGia": "GVS",
        "productLine": "",
        "diamonds": [
          {
            "giaReports": [
              "789"
            ],
            "giaReportsUrls": [
              "https://ronberto.com"
            ],
            "_id": "61813e7282a9134e32dbdc95",
            "quantity": 1,
            "carat": "1/567",
            "color": "D",
            "clarity": "FL",
            "shape": "princess cut",
            "cut": "Excellent",
            "polish": "something"
          }
        ],
        "materialDescription": "something",
        "size": "14",
        "brilliants": "0.25V"
      },
      "quotaRegime": false,
      "active": true,
      "status": "previous",
      "brand": "Roberto Coin",
      "boutiques": [
        {
          "quantity": 1,
          "_id": "60aba936d80faa257a200eb7",
          "store": "5f3a4225ffe375404f72fb06",
          "storeName": "Belgrade",
          "price": 3840,
          "priceLocal": 461000,
          "VATpercent": 20,
          "priceHistory": [
            {
              "_id": "60aba936d80faa257a200eb8",
              "date": "2021-05-21T00:00:00.000Z",
              "price": 3840,
              "VAT": 20,
              "priceLocal": 461000
            }
          ],
          "serialNumbers": [
            {
              "warrantyConfirmed": false,
              "modified": false,
              "_id": "60aba936d80faa257a200eb9",
              "number": "10811",
              "pgpReference": "10811",
              "status": "Stock",
              "location": "Drawers Sales Floor - MBD",
              "stockDate": "2021-05-20T00:00:00.000Z",
              "comment": "/",
              "reservedFor": null,
              "reservationTime": null,
              "previousStatus": "Stock"
            }
          ]
        }
      ],
      "__v": 0,
      "createdAt": "2021-05-24T13:25:10.785Z",
      "updatedAt": "2021-11-02T13:34:42.560Z"
    },
    "newActivity": {
      "manuallyAdded": false,
      "_id": "61813e7282a9134e32dbdc98",
      "type": "Product",
      "user": "6156e14ffa2cb00a0ece3403",
      "client": null,
      "comment": "Changed diamond cut from 'Poor' to 'Excellent'.",
      "wishlist": null,
      "product": "60aba936d80faa257a200eb6",
      "serialNumber": "10811",
      "createdAt": "2021-11-02T13:34:42.734Z",
      "__v": 0
    }
  }
 * @apiUse CredentialsError
 * @apiUse MissingParamsError
 * @apiUse InvalidValue
 * @apiUse InvalidJewelryType
 * @apiUse InvalidStoneType
 * @apiUse InvalidColor
 * @apiUse InvalidClarity
 * @apiUse InvalidShape
 * @apiUse InvalidCut
 * @apiUse NotFound
 */
module.exports.editProduct = async (req, res) => {
  const { productId } = req.params;
  const { brand, rmc, jewelryType, materialDescription, size, weight, stonesQty, allStonesWeight, brilliants, diaGia, purchasePrice, materials, stones, diamonds } = req.body;
  const { _id: userId } = req.user;

  // Check if all required parameters have been sent
  if (!brand) throw new Error(error.MISSING_PARAMETERS);

  // Check if 'brand' is valid
  if (!['Messika', 'Roberto Coin', 'Petrovic Diamonds'].includes(brand)) throw new Error(error.INVALID_VALUE);

  // Validate 'jewelryType'
  if (jewelryType && !jewelryTypes.includes(jewelryType)) throw new Error(error.INVALID_JEWELRY_TYPE);

  // Validate 'materials' array
  if (materials && !Array.isArray(materials)) throw new Error(error.INVALID_VALUE);

  // Validate 'stones' array
  if (stones && !Array.isArray(stones)) throw new Error(error.INVALID_VALUE);
  if (stones && stones.length && !stones.every(el => stoneTypes.includes(el.type))) throw new Error(error.INVALID_STONE_TYPE);

  // Correct diamond property values
  if (diamonds && diamonds.length) diamonds.map(el => {
    el.color = el.color.toUpperCase();
    el.clarity = el.clarity.toUpperCase();
    el.shape = el.shape.toLowerCase();
    el.cut = el.cut.toLowerCase();

    return el;
  });

  // Validate 'diamonds' array
  if (diamonds && !Array.isArray(diamonds)) throw new Error(error.INVALID_VALUE);
  if (diamonds && diamonds.length && !diamonds.every(el => colors.includes(el.color))) throw new Error(error.INVALID_COLOR);
  if (diamonds && diamonds.length && !diamonds.every(el => clarities.includes(el.clarity))) throw new Error(error.INVALID_CLARITY);
  if (diamonds && diamonds.length && !diamonds.every(el => shapes.includes(el.shape))) throw new Error(error.INVALID_SHAPE);
  if (diamonds && diamonds.length && !diamonds.every(el => cuts.includes(el.cut))) throw new Error(error.INVALID_CUT);

  // Find product
  const product = await Product.findOne({ _id: productId, brand }).lean();

  // Check if product was found
  if (!product) throw new Error(error.NOT_FOUND);

  // Create 'updateSet' object
  let updateSet = { basicInfo: { ...product.basicInfo } };

  if (rmc) updateSet.basicInfo.rmc = rmc;
  if (jewelryType) updateSet.basicInfo.jewelryType = jewelryType;
  if (materialDescription) updateSet.basicInfo.materialDescription = materialDescription;
  if (materialDescription === '') updateSet.basicInfo.materialDescription = materialDescription;
  if (size) updateSet.basicInfo.size = size;
  if (size === '') updateSet.basicInfo.size = size;
  if (weight) updateSet.basicInfo.weight = weight;
  if (weight === '') updateSet.basicInfo.weight = weight;
  if (stonesQty) updateSet.basicInfo.stonesQty = stonesQty;
  if (stonesQty === 0) updateSet.basicInfo.stonesQty = stonesQty;
  if (allStonesWeight) updateSet.basicInfo.allStonesWeight = allStonesWeight;
  if (allStonesWeight === 0) updateSet.basicInfo.allStonesWeight = allStonesWeight;
  if (brilliants) updateSet.basicInfo.brilliants = brilliants;
  if (brilliants === '') updateSet.basicInfo.brilliants = brilliants;
  if (diaGia) updateSet.basicInfo.diaGia = diaGia;
  if (diaGia === '') updateSet.basicInfo.diaGia = diaGia;
  if (purchasePrice) updateSet.basicInfo.purchasePrice = purchasePrice;
  if (purchasePrice === 0) updateSet.basicInfo.purchasePrice = purchasePrice;
  if (materials) updateSet.basicInfo.materials = materials;
  if (stones) updateSet.basicInfo.stones = stones;
  if (diamonds) updateSet.basicInfo.diamonds = diamonds;

  // Remove excessive properties not belonging to specific brand
  if (brand === 'Messika') {
    delete updateSet.basicInfo.diaGia;
    delete updateSet.basicInfo.stones;
    delete updateSet.basicInfo.diamonds;
  } else if (brand === 'Roberto Coin') {
    delete updateSet.basicInfo.materialDescription;
  } else if (brand === 'Petrovic Diamonds') {
    delete updateSet.basicInfo.materialDescription;
    delete updateSet.basicInfo.brilliants;
    delete updateSet.basicInfo.diaGia;
    delete updateSet.basicInfo.purchasePrice;
    delete updateSet.basicInfo.materials;
    delete updateSet.basicInfo.stones;
  }

  // Update product
  const updatedProduct = await Product.findOneAndUpdate(
    { _id: productId },
    { $set: updateSet },
    { new: true }
  ).lean();

  // Create 'changedFields' array
  const changedFields = [];

  // 1. Detect changes -> string properties
  if (updatedProduct.basicInfo.rmc !== product.basicInfo.rmc) changedFields.push(`Changed rmc from '${product.basicInfo.rmc}' to '${updatedProduct.basicInfo.rmc}'.`);
  if (updatedProduct.basicInfo.jewelryType !== product.basicInfo.jewelryType) changedFields.push(`Changed jewelry type from '${product.basicInfo.jewelryType}' to '${updatedProduct.basicInfo.jewelryType}'.`);

  if (updatedProduct.basicInfo.materialDescription && product.basicInfo.materialDescription && updatedProduct.basicInfo.materialDescription !== product.basicInfo.materialDescription) changedFields.push(`Changed material description from '${product.basicInfo.materialDescription}' to '${updatedProduct.basicInfo.materialDescription}'.`);
  if (updatedProduct.basicInfo.materialDescription && !product.basicInfo.materialDescription) changedFields.push(`Set material description to '${updatedProduct.basicInfo.materialDescription}'.`);
  if (!updatedProduct.basicInfo.materialDescription && product.basicInfo.materialDescription) changedFields.push(`Set material description to ''.`);

  if (updatedProduct.basicInfo.size && product.basicInfo.size && updatedProduct.basicInfo.size !== product.basicInfo.size) changedFields.push(`Changed size from '${product.basicInfo.size}' to '${updatedProduct.basicInfo.size}'.`);
  if (updatedProduct.basicInfo.size && !product.basicInfo.size) changedFields.push(`Set size to '${updatedProduct.basicInfo.size}'.`);
  if (!updatedProduct.basicInfo.size && product.basicInfo.size) changedFields.push(`Set size to ''.`);

  if (updatedProduct.basicInfo.weight && product.basicInfo.weight && updatedProduct.basicInfo.weight !== product.basicInfo.weight) changedFields.push(`Changed weight from '${product.basicInfo.weight}' to '${updatedProduct.basicInfo.weight}'.`);
  if (updatedProduct.basicInfo.weight && !product.basicInfo.weight) changedFields.push(`Set weight to '${updatedProduct.basicInfo.weight}'.`);
  if (!updatedProduct.basicInfo.weight && product.basicInfo.weight) changedFields.push(`Set weight to ''.`);

  if (updatedProduct.basicInfo.brilliants && product.basicInfo.brilliants && updatedProduct.basicInfo.brilliants !== product.basicInfo.brilliants) changedFields.push(`Changed brilliants from '${product.basicInfo.brilliants}' to '${updatedProduct.basicInfo.brilliants}'.`);
  if (updatedProduct.basicInfo.brilliants && !product.basicInfo.brilliants) changedFields.push(`Set brilliants to '${updatedProduct.basicInfo.brilliants}'.`);
  if (!updatedProduct.basicInfo.brilliants && product.basicInfo.brilliants) changedFields.push(`Set brilliants to ''.`);

  if (updatedProduct.basicInfo.diaGia && product.basicInfo.diaGia && updatedProduct.basicInfo.diaGia !== product.basicInfo.diaGia) changedFields.push(`Changed diaGia from '${product.basicInfo.diaGia}' to '${updatedProduct.basicInfo.diaGia}'.`);
  if (updatedProduct.basicInfo.diaGia && !product.basicInfo.diaGia) changedFields.push(`Set diaGia to '${updatedProduct.basicInfo.diaGia}'.`);
  if (!updatedProduct.basicInfo.diaGia && product.basicInfo.diaGia) changedFields.push(`Set diaGia to ''.`);

  // 2. Detect changes -> integer properties
  if (updatedProduct.basicInfo.stonesQty && product.basicInfo.stonesQty && updatedProduct.basicInfo.stonesQty !== product.basicInfo.stonesQty) changedFields.push(`Changed stones quantity from '${product.basicInfo.stonesQty}' to '${updatedProduct.basicInfo.stonesQty}'.`);
  if (updatedProduct.basicInfo.stonesQty && !product.basicInfo.stonesQty) changedFields.push(`Set stones quantity to '${updatedProduct.basicInfo.stonesQty}'.`);
  if (!updatedProduct.basicInfo.stonesQty && product.basicInfo.stonesQty) changedFields.push(`Set stones quantity to '0'.`);

  if (updatedProduct.basicInfo.allStonesWeight && product.basicInfo.allStonesWeight && updatedProduct.basicInfo.allStonesWeight !== product.basicInfo.allStonesWeight) changedFields.push(`Changed total weight of stones from '${product.basicInfo.allStonesWeight}' to '${updatedProduct.basicInfo.allStonesWeight}'.`);
  if (updatedProduct.basicInfo.allStonesWeight && !product.basicInfo.allStonesWeight) changedFields.push(`Set total weight of stones to '${updatedProduct.basicInfo.allStonesWeight}'.`);
  if (!updatedProduct.basicInfo.allStonesWeight && product.basicInfo.allStonesWeight) changedFields.push(`Set total weight of stones to '0'.`);

  if (updatedProduct.basicInfo.purchasePrice && product.basicInfo.purchasePrice && updatedProduct.basicInfo.purchasePrice !== product.basicInfo.purchasePrice) changedFields.push(`Changed purchase price from '${product.basicInfo.purchasePrice}' to '${updatedProduct.basicInfo.purchasePrice}'.`);
  if (updatedProduct.basicInfo.purchasePrice && !product.basicInfo.purchasePrice) changedFields.push(`Set purchase price to '${updatedProduct.basicInfo.purchasePrice}'.`);
  if (!updatedProduct.basicInfo.purchasePrice && product.basicInfo.purchasePrice) changedFields.push(`Set purchase price to '0'.`);

  // 3. Detect change in 'materials' array
  if (updatedProduct.basicInfo.materials && product.basicInfo.materials) {
    if (updatedProduct.basicInfo.materials.length > product.basicInfo.materials.length) {
      const number = updatedProduct.basicInfo.materials.length - product.basicInfo.materials.length;
      for (let i = 1; i < number + 1; i++) changedFields.push(`Added new material '${updatedProduct.basicInfo.materials[updatedProduct.basicInfo.materials.length - i]}' to materials array.`);
    } else if (updatedProduct.basicInfo.materials.length < product.basicInfo.materials.length) {
      const number = product.basicInfo.materials.length - updatedProduct.basicInfo.materials.length;
      const apostrophe = number === 1 ? '' : 's';
      changedFields.push(`Removed ${number} material${apostrophe} from materials array.`);
    } else {
      for (let i = 0; i < updatedProduct.basicInfo.materials.length; i++) {
        if (updatedProduct.basicInfo.materials.length && updatedProduct.basicInfo.materials[i] !== product.basicInfo.materials[i]) changedFields.push(`Changed material from '${product.basicInfo.materials[i]}' to '${updatedProduct.basicInfo.materials[i]}'.`);
      }
    }
  }
  if (updatedProduct.basicInfo.materials && !product.basicInfo.materials) {
    const number = updatedProduct.basicInfo.materials.length;
    for (let i = 1; i < number + 1; i++) changedFields.push(`Added new material '${updatedProduct.basicInfo.materials[updatedProduct.basicInfo.materials.length - i]}' to materials array.`);
  }

  // 4. Detect changes in 'stones' array
  if (updatedProduct.basicInfo.stones && product.basicInfo.stones) {
    if (updatedProduct.basicInfo.stones.length > product.basicInfo.stones.length) {
      const number = updatedProduct.basicInfo.stones.length - product.basicInfo.stones.length;

      for (let i = 1; i < number + 1; i++) changedFields.push(`Added new stone '${updatedProduct.basicInfo.stones[updatedProduct.basicInfo.stones.length - i].type}' to stones array.`);
    } else if (updatedProduct.basicInfo.stones.length < product.basicInfo.stones.length) {
      const number = product.basicInfo.stones.length - updatedProduct.basicInfo.stones.length;
      const apostrophe = number === 1 ? '' : 's';
      changedFields.push(`Removed ${number} stone${apostrophe} from stones array.`);
    } else {
      for (let i = 0; i < updatedProduct.basicInfo.stones.length; i++) {
        if (updatedProduct.basicInfo.stones[i].type && product.basicInfo.stones[i].type && updatedProduct.basicInfo.stones[i].type !== product.basicInfo.stones[i].type) changedFields.push(`Changed stone type from '${product.basicInfo.stones[i].type}' to '${updatedProduct.basicInfo.stones[i].type}' in stones array.`);
        if (updatedProduct.basicInfo.stones[i].type && !product.basicInfo.stones[i].type) changedFields.push(`Set stone type to '${updatedProduct.basicInfo.stones[i].type}'.`);
        if (!updatedProduct.basicInfo.stones[i].type && product.basicInfo.stones[i].type) changedFields.push(`Set stone type to ''.`);

        if (updatedProduct.basicInfo.stones[i].quantity && product.basicInfo.stones[i].quantity && updatedProduct.basicInfo.stones[i].quantity !== product.basicInfo.stones[i].quantity) changedFields.push(`Changed stone quantity from '${product.basicInfo.stones[i].quantity}' to '${updatedProduct.basicInfo.stones[i].quantity}'.`);
        if (updatedProduct.basicInfo.stones[i].quantity && !product.basicInfo.stones[i].quantity) changedFields.push(`Set stone quantity to '${updatedProduct.basicInfo.stones[i].quantity}'.`);
        if (!updatedProduct.basicInfo.stones[i].quantity && product.basicInfo.stones[i].quantity) changedFields.push(`Set stone quantity to '0'.`);

        if (updatedProduct.basicInfo.stones[i].stoneTypeWeight && product.basicInfo.stones[i].stoneTypeWeight && updatedProduct.basicInfo.stones[i].stoneTypeWeight !== product.basicInfo.stones[i].stoneTypeWeight) changedFields.push(`Changed stone type weight from '${product.basicInfo.stones[i].stoneTypeWeight}' to '${updatedProduct.basicInfo.stones[i].stoneTypeWeight}' for the stone '${updatedProduct.basicInfo.stones[i].type}.`);
        if (updatedProduct.basicInfo.stones[i].stoneTypeWeight && !product.basicInfo.stones[i].stoneTypeWeight) changedFields.push(`Set stone stone type weight to '${updatedProduct.basicInfo.stones[i].stoneTypeWeight}'.`);
        if (!updatedProduct.basicInfo.stones[i].stoneTypeWeight && product.basicInfo.stones[i].stoneTypeWeight) changedFields.push(`Set stone stone type weight to '0'.`);
      }
    }
  }
  if (updatedProduct.basicInfo.stones && !product.basicInfo.stones) {
    const number = updatedProduct.basicInfo.stones.length;
    for (let i = 1; i < number + 1; i++) changedFields.push(`Added new stone '${updatedProduct.basicInfo.stones[updatedProduct.basicInfo.stones.length - i].type}' to stones array.`);
  }

  // 5. Detect changes in 'diamonds' array
  if (updatedProduct.basicInfo.diamonds && product.basicInfo.diamonds) {
    if (updatedProduct.basicInfo.diamonds.length > product.basicInfo.diamonds.length) {
      const number = updatedProduct.basicInfo.diamonds.length - product.basicInfo.diamonds.length;

      for (let i = 1; i < number + 1; i++) changedFields.push(`Added new diamond with carat '${updatedProduct.basicInfo.diamonds[updatedProduct.basicInfo.diamonds.length - i].carat}' to diamonds array.`);
    } else if (updatedProduct.basicInfo.diamonds.length < product.basicInfo.diamonds.length) {
      const number = product.basicInfo.diamonds.length - updatedProduct.basicInfo.diamonds.length;
      const apostrophe = number === 1 ? '' : 's';
      changedFields.push(`Removed ${number} diamond${apostrophe} from diamonds array.`);
    } else {
      for (let i = 0; i < updatedProduct.basicInfo.diamonds.length; i++) {
        if (updatedProduct.basicInfo.diamonds[i].quantity && product.basicInfo.diamonds[i].quantity && updatedProduct.basicInfo.diamonds[i].quantity !== product.basicInfo.diamonds[i].quantity) changedFields.push(`Changed diamond quantity from '${product.basicInfo.diamonds[i].quantity}' to '${updatedProduct.basicInfo.diamonds[i].quantity}'.`);
        if (updatedProduct.basicInfo.diamonds[i].quantity && !product.basicInfo.diamonds[i].quantity) changedFields.push(`Set diamond quantity to '${updatedProduct.basicInfo.diamonds[i].quantity}'.`);
        if (!updatedProduct.basicInfo.diamonds[i].quantity && product.basicInfo.diamonds[i].quantity) changedFields.push(`Set diamond quantity to '0'.`);

        if (updatedProduct.basicInfo.diamonds[i].carat && product.basicInfo.diamonds[i].carat && updatedProduct.basicInfo.diamonds[i].carat !== product.basicInfo.diamonds[i].carat) changedFields.push(`Changed diamond carat from '${product.basicInfo.diamonds[i].carat}' to '${updatedProduct.basicInfo.diamonds[i].carat}'.`);
        if (updatedProduct.basicInfo.diamonds[i].carat && !product.basicInfo.diamonds[i].carat) changedFields.push(`Set diamond carat to '${updatedProduct.basicInfo.diamonds[i].carat}'.`);
        if (!updatedProduct.basicInfo.diamonds[i].carat && product.basicInfo.diamonds[i].carat) changedFields.push(`Set diamond carat to '0'.`);

        if (updatedProduct.basicInfo.diamonds[i].color && product.basicInfo.diamonds[i].color && updatedProduct.basicInfo.diamonds[i].color !== product.basicInfo.diamonds[i].color) changedFields.push(`Changed diamond color from '${product.basicInfo.diamonds[i].color}' to '${updatedProduct.basicInfo.diamonds[i].color}'.`);
        if (updatedProduct.basicInfo.diamonds[i].color && !product.basicInfo.diamonds[i].color) changedFields.push(`Set diamond color to '${updatedProduct.basicInfo.diamonds[i].color}'.`);
        if (!updatedProduct.basicInfo.diamonds[i].color && product.basicInfo.diamonds[i].color) changedFields.push(`Set diamond color to ''.`);

        if (updatedProduct.basicInfo.diamonds[i].clarity && product.basicInfo.diamonds[i].clarity && updatedProduct.basicInfo.diamonds[i].clarity !== product.basicInfo.diamonds[i].clarity) changedFields.push(`Changed diamond clarity from '${product.basicInfo.diamonds[i].clarity}' to '${updatedProduct.basicInfo.diamonds[i].clarity}'.`);
        if (updatedProduct.basicInfo.diamonds[i].clarity && !product.basicInfo.diamonds[i].clarity) changedFields.push(`Set diamond clarity to '${updatedProduct.basicInfo.diamonds[i].clarity}'.`);
        if (!updatedProduct.basicInfo.diamonds[i].clarity && product.basicInfo.diamonds[i].clarity) changedFields.push(`Set diamond clarity to ''.`);

        if (updatedProduct.basicInfo.diamonds[i].shape && product.basicInfo.diamonds[i].shape && updatedProduct.basicInfo.diamonds[i].shape !== product.basicInfo.diamonds[i].shape) changedFields.push(`Changed diamond shape from '${product.basicInfo.diamonds[i].shape}' to '${updatedProduct.basicInfo.diamonds[i].shape}'.`);
        if (updatedProduct.basicInfo.diamonds[i].shape && !product.basicInfo.diamonds[i].shape) changedFields.push(`Set diamond shape to '${updatedProduct.basicInfo.diamonds[i].shape}'.`);
        if (!updatedProduct.basicInfo.diamonds[i].shape && product.basicInfo.diamonds[i].shape) changedFields.push(`Set diamond shape to ''.`);

        if (updatedProduct.basicInfo.diamonds[i].cut && product.basicInfo.diamonds[i].cut && updatedProduct.basicInfo.diamonds[i].cut !== product.basicInfo.diamonds[i].cut) changedFields.push(`Changed diamond cut from '${product.basicInfo.diamonds[i].cut}' to '${updatedProduct.basicInfo.diamonds[i].cut}'.`);
        if (updatedProduct.basicInfo.diamonds[i].cut && !product.basicInfo.diamonds[i].cut) changedFields.push(`Set diamond cut to '${updatedProduct.basicInfo.diamonds[i].cut}'.`);
        if (!updatedProduct.basicInfo.diamonds[i].cut && product.basicInfo.diamonds[i].cut) changedFields.push(`Set diamond cut to ''.`);

        if (updatedProduct.basicInfo.diamonds[i].polish && product.basicInfo.diamonds[i].polish && updatedProduct.basicInfo.diamonds[i].polish !== product.basicInfo.diamonds[i].polish) changedFields.push(`Changed diamond polish from '${product.basicInfo.diamonds[i].polish}' to '${updatedProduct.basicInfo.diamonds[i].polish}'.`);
        if (updatedProduct.basicInfo.diamonds[i].polish && !product.basicInfo.diamonds[i].polish) changedFields.push(`Set diamond polish to '${updatedProduct.basicInfo.diamonds[i].polish}'.`);
        if (!updatedProduct.basicInfo.diamonds[i].polish && product.basicInfo.diamonds[i].polish) changedFields.push(`Set diamond polish to ''.`);

        if (updatedProduct.basicInfo.diamonds[i].symmetry && product.basicInfo.diamonds[i].symmetry && updatedProduct.basicInfo.diamonds[i].symmetry !== product.basicInfo.diamonds[i].symmetry) changedFields.push(`Changed diamond symmetry from '${product.basicInfo.diamonds[i].symmetry}' to '${updatedProduct.basicInfo.diamonds[i].symmetry}'.`);
        if (updatedProduct.basicInfo.diamonds[i].symmetry && !product.basicInfo.diamonds[i].symmetry) changedFields.push(`Set diamond symmetry to '${updatedProduct.basicInfo.diamonds[i].symmetry}'.`);
        if (!updatedProduct.basicInfo.diamonds[i].symmetry && product.basicInfo.diamonds[i].symmetry) changedFields.push(`Set diamond symmetry to ''.`);

        // 'giaReports'
        if (updatedProduct.basicInfo.diamonds[i].giaReports && product.basicInfo.diamonds[i].giaReports) {
          if (updatedProduct.basicInfo.diamonds[i].giaReports.length > product.basicInfo.diamonds[i].giaReports.length) {
            const numberGiaR = updatedProduct.basicInfo.diamonds[i].giaReports.length - product.basicInfo.diamonds[i].giaReports.length;
            for (let j = 1; j < numberGiaR + 1; j++) changedFields.push(`Added new Gia Report '${updatedProduct.basicInfo.diamonds[i].giaReports[updatedProduct.basicInfo.diamonds[i].giaReports.length - j]}' to diamonds array`);
          } else if (updatedProduct.basicInfo.diamonds[i].giaReports.length < product.basicInfo.diamonds[i].giaReports.length) {
            const numberGiaR = product.basicInfo.diamonds[i].giaReports.length - updatedProduct.basicInfo.diamonds[i].giaReports.length;
            const apostrophe = numberGiaR === 1 ? '' : 's';
            changedFields.push(`Removed ${numberGiaR} Gia Report${apostrophe} from diamonds array.`);
          } else {
            for (let j = 0; j < updatedProduct.basicInfo.diamonds[i].giaReports.length; j++) {
              if (updatedProduct.basicInfo.diamonds[i].giaReports[j] !== product.basicInfo.diamonds[i].giaReports[j]) changedFields.push(`Changed Gia Report from '${product.basicInfo.diamonds[i].giaReports[j]}' to '${updatedProduct.basicInfo.diamonds[i].giaReports[j]}'.`);
            }
          }
        }
        if (updatedProduct.basicInfo.diamonds[i].giaReports && !product.basicInfo.diamonds[i].giaReports) {
          const number = updatedProduct.basicInfo.diamonds[i].giaReports.length;
          for (let i = 1; i < number + 1; i++) changedFields.push(`Added new Gia report '${updatedProduct.basicInfo.diamonds[i].giaReports[updatedProduct.basicInfo.diamonds[i].giaReports.length - i]}' to diamond giaReports array.`);
        }

        // 'giaReportsUrls'
        if (updatedProduct.basicInfo.diamonds[i].giaReportsUrls && product.basicInfo.diamonds[i].giaReportsUrls) {
          if (updatedProduct.basicInfo.diamonds[i].giaReportsUrls.length > product.basicInfo.diamonds[i].giaReportsUrls.length) {
            const numberGiaR = updatedProduct.basicInfo.diamonds[i].giaReportsUrls.length - product.basicInfo.diamonds[i].giaReportsUrls.length;
            for (let j = 1; j < numberGiaR + 1; j++) changedFields.push(`Added new Gia Report Url'${updatedProduct.basicInfo.diamonds[i].giaReportsUrls[updatedProduct.basicInfo.diamonds[i].giaReportsUrls.length - j]}' to diamonds array`);
          } else if (updatedProduct.basicInfo.diamonds[i].giaReportsUrls.length < product.basicInfo.diamonds[i].giaReportsUrls.length) {
            const numberGiaR = product.basicInfo.diamonds[i].giaReportsUrls.length - updatedProduct.basicInfo.diamonds[i].giaReportsUrls.length;
            const apostrophe = numberGiaR === 1 ? '' : 's';
            changedFields.push(`Removed ${numberGiaR} Gia Report${apostrophe} Url${apostrophe} from diamonds array.`);
          } else {
            for (let j = 0; j < updatedProduct.basicInfo.diamonds[i].giaReportsUrls.length; j++) {
              if (updatedProduct.basicInfo.diamonds[i].giaReportsUrls[j] !== product.basicInfo.diamonds[i].giaReportsUrls[j]) changedFields.push(`Changed Gia Report Url from '${product.basicInfo.diamonds[i].giaReportsUrls[j]}' to '${updatedProduct.basicInfo.diamonds[i].giaReportsUrls[j]}'.`);
            }
          }
        }
        if (updatedProduct.basicInfo.diamonds[i].giaReportsUrls && !product.basicInfo.diamonds[i].giaReportsUrls) {
          const number = updatedProduct.basicInfo.diamonds[i].giaReportsUrls.length;
          for (let i = 1; i < number + 1; i++) changedFields.push(`Added new Gia report url '${updatedProduct.basicInfo.diamonds[i].giaReportsUrls[updatedProduct.basicInfo.diamonds[i].giaReportsUrls.length - i]}' to diamond giaReportsUrls array.`);
        }
      }
    }
  }
  if (updatedProduct.basicInfo.diamonds && !product.basicInfo.diamonds) {
    const number = updatedProduct.basicInfo.diamonds.length;
    for (let i = 1; i < number + 1; i++) changedFields.push(`Added new diamond with carat '${updatedProduct.basicInfo.diamonds[updatedProduct.basicInfo.diamonds.length - i].carat}' to diamonds array.`);
  }

  // Create 'toExecute' array
  const toExecute = [];

  // If there were any changes include those in the log comment
  const logComment = (changedFields.length > 0) ? `${changedFields.join(' ')}` : '';

  let newActivity = {};

  // Create new activity -> for watch log history
  if (changedFields.length > 0) {
    newActivity = createActivity('Product', userId, null, logComment, null, productId, updatedProduct.boutiques[0].serialNumbers[0].number);
    toExecute.push(newActivity.save());
  }

  // Execute
  await Promise.all(toExecute);

  return res.status(200).send({
    message: 'Successfully updated product',
    results: updatedProduct,
    newActivity
  });
};

/**
 * @api {get} /product/list Get product list
 * @apiVersion 1.0.0
 * @apiName getProductsList
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String='Belgrade', 'Budapest', 'Porto Montenegro'} [boutique] Filter by Boutique (Store ID)
 * @apiParam (query) {String='Rolex'} [brand] Filter by Brand Type
 * @apiParam (query) {String='CELLINI', 'OYSTER', 'PROFESSIONAL'} [collection] Filter by Collection Type
 * @apiParam (query) {String} [productLine] Filter by Product Line
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned all products",
   "results": [
     {
       "_id": "5f50e73c1e5c985976250b24",
       "status": "new",
       "brand": "Rolex",
       "boutiques": [
         {
           "quantity": 2,
           "_id": "5f5b1b6d6f4bda8815cf5484",
           "store": "5f5b1b6d6f4bda8815cf5480",
           "storeName": "Belgrade",
           "price": 1440200,
           "VATpercent": 20,
           "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
         },
         {
           "quantity": 4,
           "_id": "5f5b1b6d6f4bda8815cf5485",
           "store": "5f5b1b6d6f4bda8815cf5481",
           "storeName": "Budapest",
           "price": 1524250,
           "VATpercent": 27,
           "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
         },
         {
           "quantity": 0,
           "_id": "5f5b1b6d6f4bda8815cf5486",
           "store": "5f5b1b6d6f4bda8815cf5482",
           "storeName": "Porto Montenegro",
           "price": 1452400,
           "VATpercent": 21,
           "serialNumbers": [
            {
              "number": "FR546SAG",
              "stockDate": "2020-07-08T14:04:49.541Z"
            }
          ]
         }
       ],
       "basicInfo": {
         "rmc": "M116769TBRJ-0002",
         "collection": "PROFESSIONAL",
         "productLine": "GMT-MASTER II",
         "saleReference": "116769TBRJ",
         "materialDescription": "PAVED W-74779BRJ",
         "dial": "PAVED W",
         "bracelet": "74779BRJ",
         "box": "EN DD EMERAUDE 60",
         "exGeneveCHF": 1565800,
         "diameter": 40,
         "photos": [
           "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m116769tbrj-0002.png?impolicy=v6-upright&imwidth=420",
         ]
       },
       "wishlist": [],
       "__v": 0,
       "createdAt": "2020-09-03T12:53:16.774Z",
       "updatedAt": "2020-09-03T12:53:16.774Z"
     }
   ]
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 */
module.exports.getProductsList = async (req, res) => {
  let { store } = req.user;
  const {
    boutique = store._id,
    brand,
    collection,
    productLine,
  } = req.query;

  if (!brand || !brandTypes.includes(brand)) throw new Error(error.MISSING_PARAMETERS);

  // Creat query object
  let query = {
    brand,
  };

  // Check if 'collection' filter has been sent
  if (collection) {
    if (!collectionTypes.includes(collection)) throw new Error(error.INVALID_VALUE);
    query['basicInfo.collection'] = collection;
  }

  // Check if 'productLine' filter has been sent
  if (productLine) query['basicInfo.productLine'] = productLine;

  const boutiqueMatch = {
    'boutiques.store': ObjectId(boutique),
  };

  // Get list of products and count
  let [listOfProducts, soonInStock] = await Promise.all([
    Product.aggregate([
      { $match: query },
      { $unwind: '$boutiques' },
      { $match: boutiqueMatch },
      {
        $project: {
          brand: 1,
          boutiques: {
            price: 1,
            serialNumbers: {
              status: 1,
              stockDate: 1,
            },
          },
          basicInfo: {
            saleReference: 1,
            productLine: 1,
            collection: 1,
            rmc: 1,
            photos: 1,
            materials: 1,
            materialDescription: 1,
            diameter: 1,
            color: 1,
            size: 1,
            forModel: 1,
            forClasp: 1
          },
        },
      },
      {
        $group: {
          _id: {
            saleReference: '$basicInfo.saleReference',
            productLine: '$basicInfo.productLine',
            collection: '$basicInfo.collection',
          },
          references: {
            $push: {
              _id: '$_id',
              brand: '$brand',
              boutiques: '$boutiques',
              rmc: '$basicInfo.rmc',
              collection: '$basicInfo.collection',
              saleReference: '$basicInfo.saleReference',
              productLine: '$basicInfo.productLine',
              photos: '$basicInfo.photos',
              materials: '$basicInfo.materials',
              materialDescription: '$basicInfo.materialDescription',
              diameter: '$basicInfo.diameter',
              color: '$basicInfo.color',
              size: '$basicInfo.size',
              forModel: '$basicInfo.forModel',
              forClasp: '$basicInfo.forClasp'
            },
          },
        },
      },
      {
        $group: {
          _id: {
            productLine: '$_id.productLine',
            collection: '$_id.collection',
          },
          references: { $push: '$references' },
        },
      },
      {
        $group: {
          _id: '$_id.collection',
          productLines: {
            $push: '$$ROOT',
          },
        },
      },
    ]),
    SoonInStock.find({ store: boutique }, { product: 1 }).lean(),
  ]);

  for (let i = 0; i < listOfProducts.length; i++) {
    const productLines = listOfProducts[i].productLines;
    for (let j = 0; j < productLines.length; j++) {
      productLines[j]._id.availableForSale = 0;
      productLines[j]._id.stock = 0;
      productLines[j]._id.soonInStock = 0;
      let references = listOfProducts[i].productLines[j].references; // References for each sale reference. Array of arrays.
      const finalReferences = [];
      for (let t = 0; t < references.length; t++) {
        let uniqueReferences = references[t];
        let serialNumbers = [];

        let soonInStockCount = 0;
        const filteredReferences = [];
        uniqueReferences.forEach((u) => {
          const soon = soonInStock.filter((s) => s.product.toString() === u._id.toString());
          soonInStock = soonInStock.filter((s) => s.product.toString() !== u._id.toString());
          soonInStockCount += soon.length;
          if (u.boutiques.serialNumbers.length) {
            serialNumbers = serialNumbers.concat(u.boutiques.serialNumbers);
          }
          u.soonInStock = soon.length;
          if (u.boutiques.serialNumbers.length || u.soonInStock > 0) {
            filteredReferences.push(u);
          }
        });

        let reference = _.sortBy(filteredReferences, (obj) => (obj.boutiques.price * -1));

        reference = reference.map((r) => {
          r.soonInStock = soonInStockCount;
          return r;
        });

        if (reference.length > 1) {
          const [first, second] = reference;
          if (first.boutiques.price === second.boutiques.price) {
            const selected = _.sortBy([first, second], 'boutiques.stockDate')[0];
            selected.boutiques.serialNumbers = serialNumbers;
            finalReferences.push(selected);
          } else {
            first.boutiques.serialNumbers = serialNumbers;
            finalReferences.push(first);
          }
        } else if (reference.length === 1) {
          reference[0].boutiques.serialNumbers = serialNumbers;
          finalReferences.push(reference[0]);
        }
      }

      for (let k = 0; k < finalReferences.length; k++) {
        if (finalReferences[k]) {
          const serialNumbers = _.sortBy(_.get(finalReferences[k], 'boutiques.serialNumbers'), 'stockDate');
          const availableForSaleCount = serialNumbers.filter((s) => s.status === 'Stock').length;
          productLines[j]._id.availableForSale += availableForSaleCount;
          productLines[j]._id.stock += serialNumbers.length;
          productLines[j]._id.soonInStock += finalReferences[k].soonInStock;
          finalReferences[k].availableForSale = availableForSaleCount;
          finalReferences[k].stock = serialNumbers.length;
        }
      }
      listOfProducts[i].productLines[j].references = _.sortBy(finalReferences, 'saleReference');
    }
  }

  function findPhoto(references) {
    // If more than one, sort by stockDate and return its photo
    const referencesSorted = _.sortBy(references, (obj) => (obj.boutiques.price * -1));
    if (referencesSorted.length > 1) {
      const [first, second] = referencesSorted;
      if (first.boutiques.price === second.boutiques.price) {
        return _.sortBy([first, second], 'boutiques.stockDate')[0].photos;
      } else {
        return first.photos;
      }
    } else if (referencesSorted.length === 1) {
      return referencesSorted[0].photos;
    } else {
      return '';
    }
  }

  let professional = {};
  let oyster = {};
  let cellini = {};
  let vintage = {};

  for (let i = 0; i < listOfProducts.length; i++) {
    listOfProducts[i].availableForSale = 0;
    listOfProducts[i].stock = 0;
    listOfProducts[i].soonInStock = 0;
    listOfProducts[i].productLines = _.sortBy(listOfProducts[i].productLines, '_id.productLine');
    listOfProducts[i].productLines = listOfProducts[i].productLines.filter((p) => (p._id.stock !== 0 || p._id.soonInStock !== 0));
    for (let j = 0; j < listOfProducts[i].productLines.length; j++) {
      listOfProducts[i].availableForSale += listOfProducts[i].productLines[j]._id.availableForSale;
      listOfProducts[i].stock += listOfProducts[i].productLines[j]._id.stock;
      listOfProducts[i].soonInStock += listOfProducts[i].productLines[j]._id.soonInStock;
      listOfProducts[i].productLines[j]._id.photoUrl = findPhoto(listOfProducts[i].productLines[j].references);
    }
    if (brand === 'Rolex') {
      if (listOfProducts[i]._id === 'PROFESSIONAL') {
        professional = listOfProducts[i];
      }
      if (listOfProducts[i]._id === 'CELLINI') {
        cellini = listOfProducts[i];
      }
      if (listOfProducts[i]._id === 'OYSTER') {
        oyster = listOfProducts[i];
      }
      if (listOfProducts[i]._id === 'VINTAGE') {
        vintage = listOfProducts[i];
      }
    }
  }

  let results = listOfProducts;

  if (brand === 'Rolex') {
    results = [];
    if (Object.keys(professional).length) {
      results.push(professional);
    }
    if (Object.keys(oyster).length) {
      results.push(oyster);
    }
    if (Object.keys(cellini).length) {
      results.push(cellini);
    }
    if (Object.keys(vintage).length) {
      results.push(vintage);
    }
    results = results.filter((r) => (r.stock !== 0 || r.soonInStock !== 0));
  } else if (brand === 'SwissKubik') {
    const order = ['STARTBOX', 'MASTERBOX SINGLE', 'MASTERBOX DOUBLE', 'MASTERBOX TRIPLE', 'MASTERBOX QUADRUPLE', 'MASTERBOX SEXTUPLE', 'MASTERBOX OCTUPLE', 'MASTERBOX 12 POSITIONS', 'SPARE ITEMS', 'TRAVELBOX', 'ACCESSORIES'];
    results = _.sortBy(results, (obj) => {
      return _.indexOf(order, obj._id);
    });
    results = results.filter((r) => r.stock !== 0 || r.soonInStock !== 0);
  } else {
    results = _.sortBy(results.filter((r) => (r.stock !== 0 || r.soonInStock !== 0)), '_id');
  }

  return res.status(200).send({
    message: 'Successfully returned all products',
    results,
  });
};

/**
 * @api {get} /product/list/filter Get products filter
 * @apiVersion 1.0.0
 * @apiName getProductsListFilter
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} [boutique] Filter by Boutique (Store) ID
 * @apiParam (query) {String='Rolex'} [brand] Filter by Brand Type
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned filter data",
   "results": [
      {
        "_id": 'OYSTER',
        "count": 1,
        "productLines": [ { "name": 'PROD LINE 2', "count": 10 } ]
      },
      {
        "_id": 'CELLINI',
        "count": 2,
        "productLines": [
          { "name": 'PROD LINE', "count": 5 },
          { "name": 'PROD LINE 2', "count": 5 }
        ]
      }
   ]
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 */
module.exports.getProductsListFilter = async (req, res) => {
  let { store } = req.user;
  const {
    boutique = store._id,
    brand,
  } = req.query;

  if (!brand || !boutique) throw new Error(error.MISSING_PARAMETERS);

  let [results, soonInStock] = await Promise.all([
    Product.aggregate([
      { $match: { brand } },
      { $unwind: '$boutiques' },
      { $match: { 'boutiques.store': ObjectId(boutique) } },
      {
        $project: {
          brand: 1,
          basicInfo: {
            saleReference: 1,
            productLine: 1,
            collection: 1,
          },
          numberOfSerialNumbers: { $size: '$boutiques.serialNumbers' },
        },
      },
      {
        $group: {
          _id: {
            productLine: '$basicInfo.productLine',
            collection: '$basicInfo.collection',
          },
          count: { $sum: '$numberOfSerialNumbers' },
          data: {
            $push: '$$ROOT',
          },
        },
      },
      {
        $group: {
          _id: '$_id.collection',
          count: { $sum: '$count' },
          productLines: {
            $push: {
              name: '$_id.productLine',
              productIds: '$data._id',
              count: '$count',
            },
          },
        },
      },
    ]),
    SoonInStock.find({ store: boutique }).lean(),
  ]);

  let professional = {};
  let oyster = {};
  let cellini = {};
  let vintage = {};
  results.forEach((r) => {
    r.productLines = _.sortBy(r.productLines, 'name');
    r.productLines.forEach((productLine) => {
      let soonInStockCount = 0;
      productLine.productIds.forEach((prod) => {
        const filtered = soonInStock.filter((s) => s.product.toString() === prod.toString());
        soonInStockCount += filtered.length;
      });
      productLine.count += soonInStockCount;
      r.count += soonInStockCount;
      delete productLine.productIds;
    });
    r.productLines = r.productLines.filter((p) => (p.count !== 0));
    if (r._id === 'PROFESSIONAL') {
      professional = r;
    }
    if (r._id === 'CELLINI') {
      cellini = r;
    }
    if (r._id === 'OYSTER') {
      oyster = r;
    }
    if (r._id === 'VINTAGE') {
      vintage = r;
    }
  });
  if (brand === 'Rolex') {
    results = [];
    if (Object.keys(professional).length) {
      results.push(professional);
    }
    if (Object.keys(oyster).length) {
      results.push(oyster);
    }
    if (Object.keys(cellini).length) {
      results.push(cellini);
    }
    if (Object.keys(vintage).length) {
      results.push(vintage);
    }
    results = results.filter((r) => r.count !== 0);
  } else {
    results = _.sortBy(results.filter((r) => r.count !== 0), '_id');
  }

  return res.status(200).send({
    message: 'Successfully returned filter data',
    results,
  });
};

/**
 * @api {get} /product/list/reference Get products by reference
 * @apiVersion 1.0.0
 * @apiName getProductsByReference
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} saleReference Sale reference
 * @apiParam (query) {Boolean} [forExport] Export data to excel
 * @apiParam (query) {String} [boutique] Filter by Boutique (Store) ID
 * @apiParam (query) {String=['default', 'pgp', 'location', 'rmc', 'ref', 'dial', 'status', 'bracelet', 'materials', 'materialDescription', 'serial', 'origin', 'stockDate']} [sort] Sort by field
 * @apiParam (query) {String=['asc','desc']} [sortType] Sort ascending or descending
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned products filtered by sale reference",
   "results": [
     {
       "prices": [
         {
           "price": 8888,
           "store": "Belgrade"
         },
         {
           "price": 9999,
           "store": "Budapest"
         },
         {
           "price": 7777,
           "store": "Porto Montenegro"
         }
       ],
       "pgp": "9834",
       "location": "Rolex safe",
       "rmc": "M116500LN-0001",
       "ref": "116500LN",
       "dial": "WHITE INDEX W",
       "status": "Reserved",
       "previousStatus": "Stock",
       "bracelet": "78590",
       "serial": "J29T4806",
       "origin": "Geneva",
       "stockDate": "2020-11-30T23:00:00.000Z",
       "photos": [
           "M79500-0007.png"
       ],
       "materials": [],
       "materialDescription": "BLACK INDEX W-95750",
       "soonInStock": true,
       "reservedFor": {},
       "reservationTime": null,
       "diameter": "36",
     }
   ]
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 */
module.exports.getProductsByReference = async (req, res) => {
  let { store } = req.user;
  const { saleReference, boutique = store._id, forExport = false, sort = 'default', sortType = 'asc' } = req.query;

  const sortFields = ['default', 'pgp', 'location', 'rmc', 'ref', 'dial', 'status', 'bracelet', 'materials', 'materialDescription', 'serial', 'origin', 'stockDate'];
  if (!saleReference) throw new Error(error.MISSING_PARAMETERS);
  if (!['asc', 'desc'].includes(sortType) || !sortFields.includes(sort)) throw new Error(error.NOT_ACCEPTABLE);

  let [results, soonInStock] = await Promise.all([
    Product.aggregate([
      { $match: { 'basicInfo.saleReference': saleReference } },
      { $unwind: '$boutiques' },
      {
        $project: {
          _id: 1,
          brand: 1,
          boutiques: {
            store: 1,
            storeName: 1,
            price: 1,
            priceLocal: 1,
            serialNumbers: {
              number: 1,
              location: 1,
              status: 1,
              previousStatus: 1,
              stockDate: 1,
              pgpReference: 1,
              reservedFor: 1,
              reservationTime: 1,
              origin: 1
            }
          },
          basicInfo: {
            saleReference: 1,
            productLine: 1,
            collection: 1,
            rmc: 1,
            dial: 1,
            bracelet: 1,
            photos: 1,
            materials: 1,
            materialDescription: 1,
            diameter: 1
          }
        }
      },
      { $unwind: '$boutiques.serialNumbers' },
      {
        $lookup: {
          from: 'clients',
          localField: 'boutiques.serialNumbers.reservedFor',
          foreignField: '_id',
          as: 'boutiques.serialNumbers.reservedFor',
        }
      },
      {
        $group: {
          _id: '$_id',
          data: { $push: '$$ROOT' },
        }
      }
    ]),
    SoonInStock.aggregate([
      {
        $match: { store: ObjectId(boutique) },
      },
      {
        $lookup: {
          from: 'clients',
          localField: 'reservedFor',
          foreignField: '_id',
          as: 'reservedFor',
        }
      },
      {
        $lookup: {
          from: 'products',
          localField: 'product',
          foreignField: '_id',
          as: 'product',
        }
      },
      { $unwind: '$product' },
      {
        $match: {
          'product.basicInfo.saleReference': saleReference,
        }
      },
      { $unwind: '$product.boutiques' },
      {
        $group: {
          _id: '$serialNumber',
          data: { $push: '$$ROOT' }
        }
      }
    ])
  ]);

  let finalResults = [];
  results.forEach((result) => {
    const prices = [];
    let productSet = [];
    result.data.forEach((prod) => {
      const final = {};
      if (!prices.filter((p) => p.store === prod.boutiques.storeName).length) {
        prices.push({ price: prod.boutiques.price, store: prod.boutiques.storeName })
      }
      if (prod.boutiques.store.toString() === boutique.toString()) {
        final._id = prod._id;
        final.pgp = prod.boutiques.serialNumbers.pgpReference;
        final.photos = prod.basicInfo.photos;
        final.location = prod.boutiques.serialNumbers.location;
        final.rmc = prod.basicInfo.rmc;
        final.ref = prod.basicInfo.saleReference;
        final.dial = prod.basicInfo.dial;
        final.status = prod.boutiques.serialNumbers.status;
        final.previousStatus = prod.boutiques.serialNumbers.previousStatus;
        final.bracelet = prod.basicInfo.bracelet;
        final.materials = prod.basicInfo.materials;
        final.materialDescription = prod.basicInfo.materialDescription;
        final.serial = prod.boutiques.serialNumbers.number;
        final.origin = prod.boutiques.serialNumbers.origin;
        final.stockDate = prod.boutiques.serialNumbers.stockDate;
        final.reservedFor = {};
        if (prod.boutiques.serialNumbers.reservedFor.length) {
          final.reservedFor = prod.boutiques.serialNumbers.reservedFor[0];
        }
        final.reservationTime = null;
        if (prod.boutiques.serialNumbers.reservationTime) {
          final.reservationTime = prod.boutiques.serialNumbers.reservationTime;
        }
        final.diameter = prod.basicInfo.diameter;
      }
      if (Object.keys(final).length > 1) {
        productSet.push(final);
      }
    });
    productSet.forEach((p) => {
      p.prices = prices;
      finalResults.push(p);
    });
  });

  finalResults = _.sortBy(finalResults, 'stockDate');

  let soonInStockArray = [];
  soonInStock.forEach((serialNumber) => {
    const prices = [];
    const prod = serialNumber.data.find((s) => (s.product.boutiques.store.toString() === boutique.toString()));
    const final = {};
    serialNumber.data.forEach((s) => {
      if (!prices.filter((p) => p.store === s.product.boutiques.storeName).length) {
        prices.push({ price: s.product.boutiques.price, store: s.product.boutiques.storeName });
      }
    });
    if (prod && prod.product && prod.product.boutiques.store.toString() === boutique.toString()) {
      final._id = prod.product._id;
      final.pgp = prod.pgpReference;
      final.photos = prod.product.basicInfo.photos;
      final.location = prod.location;
      final.rmc = prod.rmc;
      final.ref = prod.product.basicInfo.saleReference;
      final.dial = prod.dial;
      final.bracelet = prod.product.basicInfo.bracelet;
      final.materials = prod.product.basicInfo.materials;
      final.materialDescription = prod.product.basicInfo.materialDescription;
      final.serial = prod.serialNumber;
      final.status = prod.status;
      final.previousStatus = prod.previousStatus;
      final.origin = prod.origin;
      final.stockDate = prod.stockDate;
      final.soonInStock = prod.soonInStock;
      final.reservedFor = {};
      if (prod.reservedFor.length) {
        final.reservedFor = prod.reservedFor[0];
      }
      final.reservationTime = null;
      if (prod.reservationTime) {
        final.reservationTime = prod.reservationTime;
      }
      final.diameter = prod.product.basicInfo.diameter;
    }
    if (Object.keys(final).length > 1) {
      final.prices = prices;
      soonInStockArray.push(final);
    }
  });
  soonInStockArray = _.sortBy(soonInStockArray, 'pgp');
  finalResults = finalResults.concat(soonInStockArray);
  if (sort !== 'default') {
    finalResults = _.orderBy(finalResults, sort, sortType);
  }

  if (forExport) {
    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('My Sheet');
    const font = { name: 'Arial', size: 12 };

    worksheet.columns = [
      {
        header: '#',
        key: 'index',
        width: 10,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: 'PGP',
        key: 'pgp',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: 'Location',
        key: 'location',
        width: 30,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: 'RMC',
        key: 'rmc',
        width: 30,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: 'Brclt',
        key: 'bracelet',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: 'Dial',
        key: 'dial',
        width: 30,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: 'Serial',
        key: 'serial',
        width: 30,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: '€ - RS',
        key: 'rsPrice',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: '€ - HU',
        key: 'huPrice',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: '€ - MN',
        key: 'mnPrice',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: 'Ref',
        key: 'ref',
        width: 20,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: 'Origin',
        key: 'origin',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
      {
        header: 'Mnt',
        key: 'stockDate',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
      },
    ];

    for (let i = 0; i < finalResults.length; i += 1) {
      const newObject = Object.assign({}, finalResults[i]);
      newObject.index = i + 1;
      if (newObject.prices.length) {
        newObject.rsPrice = _.get(newObject.prices.find((p) => p.store === 'Belgrade'), 'price');
        newObject.huPrice = _.get(newObject.prices.find((p) => p.store === 'Budapest'), 'price');
        newObject.mnPrice = _.get(newObject.prices.find((p) => p.store === 'Porto Montenegro'), 'price');
      }
      let months = '';
      if (newObject.stockDate) {
        months = moment().diff(moment(newObject.stockDate), 'months');
      }
      newObject.stockDate = months;

      worksheet.addRow(newObject);
    }

    worksheet.getRow(1).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };      // WHITE
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '000' },
    };
    worksheet.getRow(1).border = {
      top: { style: 'thin', color: { argb: 'FFFFFF' } },
      left: { style: 'thin', color: { argb: 'FFFFFF' } },
      bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
      right: { style: 'thin', color: { argb: 'FFFFFF' } },
    };
    worksheet.getRow(1).height = 30;

    const createDir = util.promisify(tmp.dir);
    const tmpDir = await createDir();
    const filePath = `${tmpDir}/${uuidv4()}.xlsx`;

    return workbook.xlsx.writeFile(filePath).then(() => {
      const stream = fs.createReadStream(filePath);

      stream.on('error', () => {
        throw new Error(error.BAD_REQUEST);
      });
      stream.pipe(res);
    });
  }

  return res.status(200).send({
    message: 'Successfully returned products filtered by sale reference',
    results: finalResults,
  });
};

/**
 * @api {get} /product/list/search Search and filter products by location
 * @apiVersion 1.0.0
 * @apiName getProductsByLocation
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} [location] Location
 * @apiParam (query) {String} [search] Search
 * @apiParam (query) {String} [boutique] Filter by Boutique (Store) ID
 * @apiParam (query) {String} [status] Filter by status
 * @apiParam (query) {String} [brand] Filter by brand
 * @apiParam (query) {Boolean} [forExport] Export data to excel
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
    "message": "Successfully returned filtered products",
    "results": [
      {
        "_id": "600f1f8099f40ffcae92ee71",
        "pgp": "99003",
        "brand": "Rolex",
        "photos": [
            "http://91.226.243.34:3029/minio/download/rolex.test/M116500LN-0001.jpg?token="
        ],
        "location": "Incoming",
        "rmc": "M116500LN-0001",
        "ref": "116500LN",
        "dial": "WHITE INDEX W",
        "bracelet": "78590",
        "serial": "600f1f8099f40ffcae92ee71",
        "status": "Stock",
        "origin": "Geneva",
        "stockDate": null,
        "productId": "603caa9776077912220b384c",
      },
      {
        "_id": "600f1f8099f40ffcae92ee72",
        "pgp": "99004",
        "brand": "Rolex",
        "photos": [
            "http://91.226.243.34:3029/minio/download/rolex.test/M116500LN-0002.jpg?token="
        ],
        "location": "Incoming",
        "rmc": "M116500LN-0002",
        "ref": "116500LN",
        "dial": "BLACK INDEX W",
        "bracelet": "78590",
        "serial": "600f1f8099f40ffcae92ee72",
        "status": "Stock",
        "origin": "Geneva",
        "stockDate": null,
        "productId": "603caa9776077912220b384c",
      },
    ]
 }
 * @apiUse MissingParamsError
 * @apiUse NotAcceptable
 * @apiUse NotFound
 */
module.exports.getProductsByLocation = async (req, res) => {
  let { store } = req.user;
  const {
    location,
    search,
    boutique = store._id,
    status,
    brand,
    forExport = false,
  } = req.query;

  if ((!location && !search && !status) || !boutique) throw new Error(error.MISSING_PARAMETERS);

  // Only one parameter is allowed, either location or search.
  if ((location && search) || (location && status) || (status && search)) throw new Error(error.NOT_ACCEPTABLE);

  // Match query for aggregate
  let matchQuery = {};

  // Match query for aggregate
  const soonInStockMatchQuery = {
    store: ObjectId(boutique)
  };
  // Second match query for SoonInStock aggregate
  const soonInStockSecondStateMatch = {
    'product.boutiques.store': ObjectId(boutique)
  };

  // Set matchQuery to search by multiple fields
  if (search) {
    soonInStockSecondStateMatch.$or = [
      { serialNumber: new RegExp(search, 'i') },
      { location: new RegExp(search, 'i') },
      { status: new RegExp(search, 'i') },
      { pgpReference: new RegExp(search, 'i') },
      { origin: new RegExp(search, 'i') },
      { rmc: new RegExp(search, 'i') },
      { dial: new RegExp(search, 'i') },
      { 'product.boutiques.price': Number(search) },
      { 'product.basicInfo.saleReference': new RegExp(search, 'i') },
      { 'product.basicInfo.productLine': new RegExp(search, 'i') },
      { 'product.basicInfo.collection': new RegExp(search, 'i') },
      { 'product.basicInfo.bracelet': new RegExp(search, 'i') },
      { 'reservedFor.nameSearch': new RegExp(search, 'i') },
      { comment: new RegExp(search, 'i') }
    ];
    matchQuery = {
      'boutiques.store': ObjectId(boutique),
      $or: [
        { 'boutiques.serialNumbers.number': new RegExp(search, 'i') },
        { 'boutiques.serialNumbers.location': new RegExp(search, 'i') },
        { 'boutiques.serialNumbers.status': new RegExp(search, 'i') },
        { 'boutiques.serialNumbers.pgpReference': new RegExp(search, 'i') },
        { 'boutiques.serialNumbers.origin': new RegExp(search, 'i') },
        { 'boutiques.price': Number(search) },
        { 'basicInfo.saleReference': new RegExp(search, 'i') },
        { 'basicInfo.productLine': new RegExp(search, 'i') },
        { 'basicInfo.collection': new RegExp(search, 'i') },
        { 'basicInfo.rmc': new RegExp(search, 'i') },
        { 'basicInfo.dial': new RegExp(search, 'i') },
        { 'basicInfo.bracelet': new RegExp(search, 'i') },
        { 'boutiques.serialNumbers.reservedFor.nameSearch': new RegExp(search, 'i') },
        { 'boutiques.serialNumbers.comment': new RegExp(search, 'i') }
      ]
    };
  }

  // Set matchQuery to filter products by location
  if (location) {
    soonInStockMatchQuery.location = new RegExp(location, 'i');
    matchQuery = {
      'boutiques.serialNumbers.location': location,
      'boutiques.store': ObjectId(boutique)
    };
  }

  // Set matchQuery to filter products by status
  if (status) {
    soonInStockMatchQuery.status = new RegExp(status, 'i');
    matchQuery = {
      'boutiques.serialNumbers.status': status,
      'boutiques.store': ObjectId(boutique)
    };
  }

  // Set matchQuery to filter products by brand
  if (brand) {
    soonInStockSecondStateMatch['product.brand'] = brand;
    matchQuery.brand = brand;
    matchQuery['boutiques.store'] = ObjectId(boutique);
  }

  // Set projection stage object
  const projectionSet = {
    _id: 1,
    brand: 1,
    boutiques: {
      store: 1,
      storeName: 1,
      price: 1,
      priceLocal: 1,
      serialNumbers: {
        number: 1,
        location: 1,
        status: 1,
        stockDate: 1,
        pgpReference: 1,
        reservedFor: 1,
        reservationTime: 1,
        origin: 1,
        comment: 1,
        warrantyConfirmed: 1,
        exGenevaPrice: 1
      }
    },
    basicInfo: {
      saleReference: 1,
      productLine: 1,
      collection: 1,
      rmc: 1,
      dial: 1,
      bracelet: 1,
      photos: 1,
      jewelryType: 1,
      size: 1,
      weight: 1,
      stonesQty: 1,
      allStonesWeight: 1,
      diamonds: 1,
      materials: 1,
      materialDescription: 1,
      diameter: 1
    }
  };

  // Aggregation for product
  const productAggregate = Product.aggregate([
    { $unwind: '$boutiques' },
    {
      $project: projectionSet
    },
    { $unwind: '$boutiques.serialNumbers' },
    {
      $lookup: {
        from: 'clients',
        localField: 'boutiques.serialNumbers.reservedFor',
        foreignField: '_id',
        as: 'boutiques.serialNumbers.reservedFor'
      },
    },
    {
      $match: matchQuery
    },
    {
      $group: {
        _id: '$boutiques.serialNumbers.number',
        product: { $first: '$$ROOT' }
      },
    },
    {
      $sort: { 'product.boutiques.serialNumbers.stockDate': 1 }
    },
  ]);
  const soonInStockAggregate = SoonInStock.aggregate([
    {
      $match: soonInStockMatchQuery
    },
    {
      $lookup: {
        from: 'clients',
        localField: 'reservedFor',
        foreignField: '_id',
        as: 'reservedFor'
      },
    },
    {
      $lookup: {
        from: 'products',
        localField: 'product',
        foreignField: '_id',
        as: 'product'
      }
    },
    {
      $unwind: {
        path: '$product',
        preserveNullAndEmptyArrays: true
      }
    },
    {
      $unwind: {
        path: '$product.boutiques',
        preserveNullAndEmptyArrays: true
      }
    },
    {
      $unwind: {
        path: '$product.boutiques.serialNumbers',
        preserveNullAndEmptyArrays: true
      }
    },
    {
      $match: soonInStockSecondStateMatch
    },
    {
      $group: {
        _id: '$serialNumber',
        product: { $first: '$$ROOT' }
      }
    },
    {
      $sort: { 'product.product.basicInfo.saleReference': 1 }
    }
  ]);

  let results = [];
  let soonInStock = [];

  if (location && (location === 'Incoming' || location.includes('Expected'))) {
    results = await soonInStockAggregate;
  } else if (location && location.length && (location !== 'Incoming' || !location.includes('Expected'))) {
    results = await productAggregate;
  } else {
    // Search set
    [results, soonInStock] = await Promise.all([
      productAggregate,
      soonInStockAggregate
    ]);
  }

  results = results.map((prod) => {
    if (location && (location === 'Incoming' || location.includes('Expected'))) {
      return {
        _id: prod.product._id,
        pgp: prod.product.pgpReference,
        brand: prod.product.product.brand,
        photos: prod.product.product.basicInfo.photos,
        location: prod.product.location,
        rmc: prod.product.product.basicInfo.rmc,
        ref: prod.product.product.basicInfo.saleReference,
        productLine: prod.product.product.basicInfo.productLine,
        collection: prod.product.product.basicInfo.collection,
        dial: prod.product.product.basicInfo.dial,
        bracelet: prod.product.product.basicInfo.bracelet,
        serial: prod.product.serialNumber,
        price: prod.product.product.boutiques.price,
        store: prod.product.product.boutiques.storeName,
        status: prod.product.status,
        warrantyConfirmed: prod.product.warrantyConfirmed,
        exGenevaPrice: prod.product.exGenevaPrice,
        origin: 'Geneva',
        stockDate: prod.product.stockDate,
        soonInStock: prod.product.soonInStock,
        productId: prod.product.product._id,
        reservedFor: {
          _id: prod.product.reservedFor.length ? prod.product.reservedFor[0]._id : '',
          fullName: prod.product.reservedFor.length ? prod.product.reservedFor[0].fullName : '',
        },
        reservationTime: prod.product.reservationTime,
        diameter: prod.product.product.basicInfo.diameter
      }
    }
    let diamondShape = '';
    let color = '';
    let carat = '';
    let certificate = '';
    let diaCt = '';
    const diamonds = _.get(prod, 'product.basicInfo.diamonds');
    if (diamonds && diamonds.length) {
      diamondShape = `${_.get(prod, 'product.basicInfo.diamonds[0].shape')} × ${_.get(prod, 'product.basicInfo.diamonds').length}`;
      color = _.get(prod, 'product.basicInfo.diamonds[0].clarity');
      carat = _.get(prod, 'product.basicInfo.diamonds[0].carat');
      certificate = diamonds[0].giaReports;
      diaCt = diamonds[0].carat;
    }
    let rubyCt = '';
    const stones = _.get(prod, 'product.basicInfo.stones');
    if (stones && stones.length) {
      rubyCt = stones[0].stoneTypeWeight;
    }
    return {
      _id: prod.product._id,
      pgp: prod.product.boutiques.serialNumbers.pgpReference,
      brand: prod.product.brand,
      photos: prod.product.basicInfo.photos,
      location: prod.product.boutiques.serialNumbers.location,
      rmc: prod.product.basicInfo.rmc,
      ref: prod.product.basicInfo.saleReference,
      productLine: prod.product.basicInfo.productLine,
      dial: prod.product.basicInfo.dial,
      bracelet: prod.product.basicInfo.bracelet,
      serial: prod._id,
      price: prod.product.boutiques.price,
      store: prod.product.boutiques.storeName,
      status: prod.product.boutiques.serialNumbers.status,
      warrantyConfirmed: prod.product.boutiques.serialNumbers.warrantyConfirmed,
      exGenevaPrice: prod.product.boutiques.serialNumbers.exGenevaPrice,
      origin: prod.product.boutiques.serialNumbers.origin,
      stockDate: prod.product.boutiques.serialNumbers.stockDate,
      soonInStock: false,
      productId: prod.product._id,
      reservedFor: {
        _id: prod.product.boutiques.serialNumbers.reservedFor.length ? prod.product.boutiques.serialNumbers.reservedFor[0]._id : '',
        fullName: prod.product.boutiques.serialNumbers.reservedFor.length ? prod.product.boutiques.serialNumbers.reservedFor[0].fullName : '',
      },
      reservationTime: prod.product.boutiques.serialNumbers.reservationTime,
      type: prod.product.basicInfo.jewelryType,
      size: prod.product.basicInfo.size,
      weight: prod.product.basicInfo.weight,
      diaQty: prod.product.basicInfo.stonesQty,
      gc: prod.product.basicInfo.materials,
      collection: prod.product.basicInfo.collection,
      goldGr: prod.product.basicInfo.weight,
      carat,
      color,
      diamondShape,
      certificate,
      diaCt,
      rubyCt,
      materialDescription: prod.product.basicInfo.materialDescription,
      materials: prod.product.basicInfo.materials,
      diameter: prod.product.basicInfo.diameter
    }
  });

  soonInStock = soonInStock.map((prod) => {
    return {
      _id: prod.product._id,
      pgp: prod.product.pgpReference,
      brand: prod.product.product.brand,
      photos: prod.product.product.basicInfo.photos,
      location: prod.product.location,
      rmc: prod.product.product.basicInfo.rmc,
      ref: prod.product.product.basicInfo.saleReference,
      productLine: prod.product.product.basicInfo.productLine,
      collection: prod.product.product.basicInfo.collection,
      dial: prod.product.product.basicInfo.dial,
      bracelet: prod.product.product.basicInfo.bracelet,
      serial: prod.product.serialNumber,
      price: prod.product.product.boutiques.price,
      store: prod.product.product.boutiques.storeName,
      status: prod.product.status,
      warrantyConfirmed: prod.product.warrantyConfirmed,
      exGenevaPrice: prod.product.exGenevaPrice,
      origin: 'Geneva',
      stockDate: prod.product.stockDate,
      soonInStock: prod.product.soonInStock,
      productId: prod.product.product._id,
      reservedFor: {
        _id: prod.product.reservedFor.length ? prod.product.reservedFor[0]._id : '',
        fullName: prod.product.reservedFor.length ? prod.product.reservedFor[0].fullName : '',
      },
      reservationTime: prod.product.reservationTime,
      diameter: prod.product.product.basicInfo.diameter
    }
  });

  results = results.concat(soonInStock);
  const count = results.length;

  if (forExport) {
    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('My Sheet');
    worksheet.addRow();
    const font = { name: 'Arial', size: 12 };

    worksheet.columns = [
      {
        header: '#',
        key: 'index',
        width: 10,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
      },
      {
        header: 'PGP',
        key: 'pgp',
        width: 10,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
      },
      {
        header: 'Location',
        key: 'location',
        width: 30,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      },
      {
        header: 'Status',
        key: 'status',
        width: 30,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      },
      {
        header: 'Reserved For',
        key: 'reserved',
        width: 30,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      },
      {
        header: 'RX Ref',
        key: 'rXRef',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      },
      {
        header: 'RMC',
        key: 'rmc',
        width: 20,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      },
      {
        header: 'Brac No',
        key: 'bracelet',
        width: 20,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      },
      {
        header: 'Warranty',
        key: 'warranty',
        width: 10,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
      },
      {
        header: 'S/N',
        key: 'serial',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
      },
      {
        header: 'Dial',
        key: 'dial',
        width: 30,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      },
      {
        header: 'Retail EUR',
        key: 'priceEur',
        width: 15,
        style: { numFmt: '#,##0;[Red]-#,##0', font, alignment: { vertical: 'middle', horizontal: 'right' } }
      },
      {
        header: 'Ex Geneve CHF',
        key: 'exGeneveCHF',
        width: 20,
        style: { numFmt: '#,##0;[Red]-#,##0', font, alignment: { vertical: 'middle', horizontal: 'right' } }
      },
      {
        header: 'Model',
        key: 'model',
        width: 30,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      },
      {
        header: 'Collection',
        key: 'collection',
        width: 20,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      },
      {
        header: 'Age',
        key: 'age',
        width: 10,
        style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
      },
      {
        header: 'Store',
        key: 'store',
        width: 15,
        style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
      }
    ];
    let count = 0;
    for (const product of results) {

      const newObject = Object.assign({}, product);
      newObject.index = count + 1;
      newObject.pgp = newObject.pgp;
      newObject.location = newObject.location;
      newObject.status = newObject.status;
      newObject.reserved =  newObject.reservedFor ? newObject.reservedFor.fullName : '';
      newObject.rXRef = newObject.ref;
      newObject.rmc = newObject.rmc;
      newObject.bracelet = newObject.bracelet;
      newObject.warranty = newObject.warrantyConfirmed ? 1 : 0;
      newObject.serial = newObject.serial;
      newObject.dial = newObject.dial;
      newObject.priceEur = newObject.price;
      newObject.exGeneveCHF = newObject.exGenevaPrice / 100;
      newObject.model = newObject.productLine;
      newObject.collection = newObject.collection;
      newObject.age = newObject.stockDate ? Math.ceil(Math.abs(new Date() - new Date(newObject.stockDate)) / (1000 * 30 * 60 * 60 * 24)) : 0;
      newObject.store = newObject.store;
      worksheet.addRow(newObject, 'i');
      count += 1;
    }

    worksheet.getRow(1).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '000' }
    };
    worksheet.getRow(1).border = {
      top: { style: 'thin', color: { argb: 'FFFFFF' } },
      left: { style: 'thin', color: { argb: 'FFFFFF' } },
      bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
      right: { style: 'thin', color: { argb: 'FFFFFF' } }
    };
    worksheet.getRow(1).height = 30;

    const createDir = util.promisify(tmp.dir);
    const tmpDir = await createDir();
    const filePath = `${tmpDir}/${uuidv4()}.xlsx`;

    return workbook.xlsx.writeFile(filePath).then(() => {
      const stream = fs.createReadStream(filePath);

      stream.on('error', () => {
        throw new Error(error.BAD_REQUEST);
      });
      stream.pipe(res);
    });
  }

  return res.status(200).send({
    message: 'Successfully returned filtered products',
    count,
    results
  });
};

/**
 * @api {get} /product/list/jewelry Get jewelry list by brand
 * @apiVersion 1.0.0
 * @apiName getJewelryList
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String=['Petrovic Diamonds', 'Roberto Coin']} brand Brand name
 * @apiParam (query) {String} [boutique] Filter by Boutique (Store) ID
 * @apiParam (query) {Boolean} [forExport] Export data to excel
 * @apiParam (query) {String=['type','collection']} [groupBy] Group by type or by collection
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 Response for brand Petrovic Diamonds:
 {
      "message": "Successfully returned jewelry list",
      "results": [
          {
              "_id": "earrings",
              "stock": 1,
              "products": [
                  {
                      "_id": "600f20dafada03fde567a29b",
                      "pgp": "5388",
                      "photos": [
                          "http://91.226.243.34:3029/minio/download/rolex.test/ERB025049.jpg?token=",
                          "http://91.226.243.34:3029/minio/download/rolex.test/ERB025049_second.jpg?token=",
                          "http://91.226.243.34:3029/minio/download/rolex.test/ERB025049_third.jpg?token="
                      ],
                      "location": "Drawer display - MBD",
                      "reservedFor": {},
                      "status": "Stock",
                      "previousStatus": "Reserved",
                      "reservationTime": "2021-07-15T14:03:50.158Z",
                      "type": "earrings",
                      "serial": "5388",
                      "size": "",
                      "weight": "1.5",
                      "carat": "0.4",
                      "color": "VVS2",
                      "diamondShape": "round brilliant × 2",
                      "price": 3170,
                      "stockDate": "2017-12-15T23:00:00.000Z"
                  }
              ]
          },
          {
              "_id": "montature",
              "stock": 1,
              "products": [
                  {
                      "_id": "600f20dafada03fde567a291",
                      "pgp": "1823",
                      "photos": [
                          "http://91.226.243.34:3029/minio/download/rolex.test/M000024.jpg?token=",
                          "http://91.226.243.34:3029/minio/download/rolex.test/M000024_second.jpg?token=",
                          "http://91.226.243.34:3029/minio/download/rolex.test/M000024_third.jpg?token="
                      ],
                      "location": "Drawer display - MBD",
                      "reservedFor": {},
                      "status": "Stock",
                      "previousStatus": "Reserved",
                      "reservationTime": "2021-07-15T14:03:50.158Z",
                      "type": "montature",
                      "serial": "1823",
                      "size": "13",
                      "weight": "5.53",
                      "carat": "0",
                      "diamondShape": "round brilliant × 1",
                      "price": 1210,
                      "stockDate": "2014-03-13T23:00:00.000Z"
                  }
              ]
          }
      ]
  }
 Response for brand Roberto Coin(grouped by collection):
 {
      "message": "Successfully returned jewelry list",
      "results": [
          {
              "_id": "Wedding Band",
              "stock": 1,
              "products": [
                  {
                      "_id": "600f20e764bf28fe05c84c19",
                      "pgp": "6412",
                      "photos": [
                          "http://91.226.243.34:3029/minio/download/rolex.test/ADR449RI0543_12.jpg?token="
                      ],
                      "location": "Drawer display - MBD",
                      "reservedFor": {},
                      "status": "Stock",
                      "previousStatus": "Reserved",
                      "reservationTime": "2021-07-15T14:03:50.158Z",
                      "saleReference": "ADR449RI0543_12",
                      "serial": "6412",
                      "type": "ring",
                      "collection": "Wedding Band",
                      "goldGr": "3.78",
                      "diaQty": 0,
                      "diaCt": "0.2",
                      "rubyCt": "",
                      "certificate": [
                          ""
                      ],
                      "price": 1650,
                      "stockDate": "2018-08-06T22:00:00.000Z"
                  }
              ]
          },
          {
              "_id": "Princess Flower",
              "stock": 1,
              "products": [
                  {
                      "_id": "600f20e864bf28fe05c84c8b",
                      "pgp": "8144",
                      "photos": [
                          "http://91.226.243.34:3029/minio/download/rolex.test/ADR777BR2662.jpg?token="
                      ],
                      "location": "Shop windows - MBD",
                      "reservedFor": {},
                      "status": "Stock",
                      "previousStatus": "Reserved",
                      "reservationTime": "2021-07-15T14:03:50.158Z",
                      "saleReference": "ADR777BR2662",
                      "serial": "8144",
                      "type": "bracelet",
                      "collection": "Princess Flower",
                      "goldGr": "11.06",
                      "diaQty": 40,
                      "diaCt": "0.444",
                      "rubyCt": "",
                      "certificate": [
                          ""
                      ],
                      "price": 6420,
                      "stockDate": "2019-06-09T22:00:00.000Z"
                  },
              ]
          }
      ]
  }
 * @apiUse MissingParamsError
 * @apiUse NotAcceptable
 * @apiUse NotFound
 */
module.exports.getJewelryList = async (req, res) => {
  let { store } = req.user;
  const {
    boutique = store._id,
    brand,
    forExport = false,
  } = req.query;
  let { groupBy = 'type' } = req.query;

  if (!boutique || !brand) throw new Error(error.MISSING_PARAMETERS);
  if (!['Petrovic Diamonds', 'Roberto Coin', 'Messika'].includes(brand) || !['type', 'collection'].includes(groupBy)) throw new Error(error.NOT_ACCEPTABLE);

  if (groupBy === 'type') {
    groupBy = '$product.basicInfo.jewelryType';
  } else {
    groupBy = '$product.basicInfo.collection';
  }

  // Aggregation for product
  let results = await Product.aggregate([
    { $unwind: '$boutiques' },
    {
      $project: {
        _id: 1,
        brand: 1,
        boutiques: {
          store: 1,
          storeName: 1,
          price: 1,
          priceLocal: 1,
          serialNumbers: {
            number: 1,
            location: 1,
            status: 1,
            previousStatus: 1,
            stockDate: 1,
            pgpReference: 1,
            reservedFor: 1,
            reservationTime: 1,
            origin: 1
          }
        },
        basicInfo: {
          saleReference: 1,
          jewelryType: 1,
          collection: 1,
          size: 1,
          weight: 1,
          stonesQty: 1,
          allStonesWeight: 1,
          diamonds: 1,
          rmc: 1,
          photos: 1,
          materials: 1
        }
      }
    },
    { $unwind: '$boutiques.serialNumbers' },
    {
      $match: {
        'boutiques.store': ObjectId(boutique),
        brand
      }
    },
    {
      $lookup: {
        from: 'clients',
        localField: 'boutiques.serialNumbers.reservedFor',
        foreignField: '_id',
        as: 'boutiques.serialNumbers.reservedFor'
      }
    },
    {
      $group: {
        _id: '$boutiques.serialNumbers.number',
        product: { $first: '$$ROOT' }
      }
    },
    {
      $group: {
        _id: groupBy,
        products: { $push: '$product' },
        stock: { $sum: 1 }
      }
    },
    { $sort: { _id: 1 } }
  ]);

  if (brand === 'Petrovic Diamonds') {
    results = results.map((collection) => ({
      _id: collection._id,
      products: _.sortBy(collection.products.map((prod) => {
        let diamondShape = '';
        let color = '';
        let carat = '';
        const diamonds = _.get(prod, 'basicInfo.diamonds');
        if (diamonds && diamonds.length) {
          diamondShape = `${_.get(prod, 'basicInfo.diamonds[0].shape')} × ${_.get(prod, 'basicInfo.diamonds').length}`;
          color = _.get(prod, 'basicInfo.diamonds[0].clarity');
          carat = _.get(prod, 'basicInfo.diamonds[0].carat');
        }
        return {
          _id: prod._id,
          pgp: prod.boutiques.serialNumbers.pgpReference,
          photos: prod.basicInfo.photos,
          location: prod.boutiques.serialNumbers.location,
          status: prod.boutiques.serialNumbers.status,
          previousStatus: prod.boutiques.serialNumbers.previousStatus,
          reservedFor: _.get(prod, 'boutiques.serialNumbers.reservedFor[0]') || {},
          reservationTime: _.get(prod, 'boutiques.serialNumbers.reservationTime') || null,
          type: prod.basicInfo.jewelryType,
          serial: prod.boutiques.serialNumbers.number,
          size: prod.basicInfo.size,
          weight: prod.basicInfo.weight,
          carat,
          color,
          diamondShape,
          price: prod.boutiques.price,
          stockDate: prod.boutiques.serialNumbers.stockDate
        }
      }), 'stockDate'),
      stock: collection.stock
    }));
  } else {
    results = results.map((collection) => {
      return {
        _id: collection._id,
        products: _.sortBy(collection.products.map((prod) => {
          let certificate = '';
          let diaCt = '';
          const diamonds = _.get(prod, 'basicInfo.diamonds');
          if (diamonds && diamonds.length) {
            certificate = diamonds[0].giaReports;
            diaCt = diamonds[0].carat;
          }

          let rubyCt = '';
          const stones = _.get(prod, 'basicInfo.stones');
          if (stones && stones.length) {
            rubyCt = stones[0].stoneTypeWeight;
          }
          return {
            _id: prod._id,
            pgp: prod.boutiques.serialNumbers.pgpReference,
            photos: prod.basicInfo.photos,
            location: prod.boutiques.serialNumbers.location,
            status: prod.boutiques.serialNumbers.status,
            previousStatus: prod.boutiques.serialNumbers.previousStatus,
            reservedFor: _.get(prod, 'boutiques.serialNumbers.reservedFor[0]') || {},
            reservationTime: _.get(prod, 'boutiques.serialNumbers.reservationTime') || null,
            saleReference: prod.basicInfo.saleReference,
            serial: prod.boutiques.serialNumbers.number,
            gc: prod.basicInfo.materials,
            type: prod.basicInfo.jewelryType,
            collection: prod.basicInfo.collection,
            goldGr: prod.basicInfo.weight,
            diaQty: prod.basicInfo.stonesQty,
            diaCt,
            rubyCt,
            certificate,
            price: prod.boutiques.price,
            stockDate: prod.boutiques.serialNumbers.stockDate
          };
        }), 'stockDate'),
        stock: collection.stock
      };
    });
  }

  if (forExport) {
    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('My Sheet');

    const font = { name: 'Arial', size: 12 };
    if (brand === 'Petrovic Diamonds') {
      worksheet.columns = [
        {
          header: '#',
          key: 'index',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'PGP ref',
          key: 'pgp',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Location',
          key: 'location',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Type',
          key: 'type',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Size',
          key: 'size',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Color Clarity',
          key: 'color',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Weight',
          key: 'weight',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Diamond shape',
          key: 'diamondShape',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: '€ - RS',
          key: 'price',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Mnt',
          key: 'stockDate',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
      ];
    } else {
      worksheet.columns = [
        {
          header: '#',
          key: 'index',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'PGP ref',
          key: 'pgp',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Location',
          key: 'location',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Type',
          key: 'type',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Collection',
          key: 'collection',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Material',
          key: 'material',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Gold gr',
          key: 'goldGr',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Dia Qty',
          key: 'diaQty',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Ruby Ct',
          key: 'rubyCt',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Roberto Ref.',
          key: 'saleReference',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Certificate',
          key: 'certificate',
          width: 25,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: '€ - RS',
          key: 'price',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        },
        {
          header: 'Mnt',
          key: 'stockDate',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
        }
      ];
    }

    let count = 0
    for (let i = 0; i < results.length; i += 1) {
      for (let j = 0; j < results[i].products.length; j += 1) {
        const newObject = Object.assign({}, results[i].products[j]);
        count += 1
        newObject.index = count;
        let months = '';
        if (newObject.stockDate) {
          months = moment().diff(moment(newObject.stockDate), 'months');
        }
        if (newObject.certificate) {
          newObject.certificate = newObject.certificate.toString();
        }
        newObject.stockDate = months;

        worksheet.addRow(newObject);
      }
    }

    worksheet.getRow(1).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '000' }
    };
    worksheet.getRow(1).border = {
      top: { style: 'thin', color: { argb: 'FFFFFF' } },
      left: { style: 'thin', color: { argb: 'FFFFFF' } },
      bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
      right: { style: 'thin', color: { argb: 'FFFFFF' } }
    };
    worksheet.getRow(1).height = 30;

    const createDir = util.promisify(tmp.dir);
    const tmpDir = await createDir();
    const filePath = `${tmpDir}/${uuidv4()}.xlsx`;

    return workbook.xlsx.writeFile(filePath).then(() => {
      const stream = fs.createReadStream(filePath);

      stream.on('error', () => {
        throw new Error(error.BAD_REQUEST);
      });
      stream.pipe(res);
    });
  }

  return res.status(200).send({
    message: 'Successfully returned jewelry list',
    results,
  });
};

/**
 * @api {get} /product/list/jewelry/filter Get jewelry filter
 * @apiVersion 1.0.0
 * @apiName getJewelryList
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String=['Petrovic Diamonds', 'Roberto Coin', 'Messika']} brand Brand name
 * @apiParam (query) {String} [boutique] Filter by Boutique (Store) ID
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
    "message": "Successfully returned jewelry list",
    "results": [
        {
            "_id": "earrings",
            "count": 12
        },
        {
            "_id": "montature",
            "count": 1
        },
        {
            "_id": "pendant",
            "count": 1
        },
        {
            "_id": "solitaire ring",
            "count": 24
        },
        {
            "_id": "trilogy ring",
            "count": 3
        }
    ]
 }
 * @apiUse MissingParamsError
 * @apiUse NotAcceptable
 * @apiUse NotFound
 */
module.exports.getJewelryListFilter = async (req, res) => {
  let { store } = req.user;
  const {
    boutique = store._id,
    brand,
  } = req.query;

  if (!boutique || !brand) throw new Error(error.MISSING_PARAMETERS);
  if (!['Petrovic Diamonds', 'Roberto Coin', 'Messika'].includes(brand)) throw new Error(error.NOT_ACCEPTABLE);

  const results = await Product.aggregate([
    { $match: { brand } },
    { $unwind: '$boutiques' },
    { $match: { 'boutiques.store': ObjectId(boutique) } },
    {
      $project: {
        brand: 1,
        basicInfo: {
          saleReference: 1,
          productLine: 1,
          collection: 1,
          jewelryType: 1
        },
        numberOfSerialNumbers: { $size: '$boutiques.serialNumbers' }
      }
    },
    {
      $group: {
        _id: '$basicInfo.jewelryType',
        count: { $sum: '$numberOfSerialNumbers' }
      }
    },
    { $sort: { _id: 1 } }
  ]);

  return res.status(200).send({
    message: 'Successfully returned jewelry list',
    results
  });
};

/**
 * @api {get} /product/list/location Get product locations
 * @apiVersion 1.0.0
 * @apiName getProductLocations
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String=true} [inStock] Filter products in stock
 * @apiParam (query) {String} [storeId] Filter by Store
 * @apiParam (query) {String=true} [multibrand] Filter Rolex or multibrand products
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned product locations",
   "count": 5,
   "results": [
     "ACC safe",
     "Awaiting CITES approval",
     "Drawer",
     "Drawer Sales Floor",
     "Drawer Sales Floor - MBD",
   ]
 }
 * @apiUse InvalidValue
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
module.exports.getProductLocations = async (req, res) => {
  const { inStock, storeId, multibrand } = req.query;

  // Creat 'query' objects
  let queryStock = {};
  let querySoonInStock = {};
  let queryStockBoutiques = {};

  // Check if 'storeId' filter has been sent
  if (storeId) {
    if (!isValidId(storeId)) throw new Error(error.INVALID_VALUE);
    queryStockBoutiques = { 'boutiques.store': ObjectId(storeId) };
    querySoonInStock = { store: storeId };
  }

  // Check if 'multibrand' filter has been sent
  let productQuery = multibrand === 'true'
      ? { $or: [{ brand: 'Tudor' }, { brand: 'Panerai' }, { brand: 'Bvlgari' }, { brand: 'Messika' }, { brand: 'Petrovic Diamonds' }, { brand: 'Roberto Coin' }, { brand: 'SwissKubik' }, { brand: 'Rubber B' }] }
      : { brand: 'Rolex' };

  // Find products, clients and stores by filter data
  const products = await Product.find(productQuery, { _id: 1 }).lean();

  const productFilterApplied = !!Object.keys(productQuery).length;

  // If any of Product is empty(no match), set find by ID to null
  if (productFilterApplied && !products.length) {
    queryStock.product = null;
    querySoonInStock.product = null;
  } else if (productFilterApplied && products.length) {
    // If filter was applied
    queryStock._id = { $in: products.map((p) => p._id) };
    querySoonInStock.product = { $in: products.map((p) => p._id) };
  }

  // Find locations for whatch in stock and soon in stock
  const [locationsStock, locationsSoonInStock] = await Promise.all([
    Product.aggregate([
      { $match: queryStock },
      { $unwind: '$boutiques' },
      { $match: queryStockBoutiques },
      { $unwind: '$boutiques.serialNumbers' },
      {
        $group: {
          _id: '$boutiques.serialNumbers.location',
        },
      },
      {
        $sort: { 'boutiques.serialNumbers.location': 1 }
      }
    ]),
    SoonInStock.distinct('location', querySoonInStock)
  ]);

  const locationsStockMap = locationsStock.map((a) => a._id);

  let location = locationsStockMap.concat(locationsSoonInStock);

  if (storeId && inStock !== 'true') {
    const store = await Store.findOne({ _id: storeId }).lean();
    if (!store) throw new Error(error.NOT_FOUND);
    switch (store.name) {
      case 'Belgrade':
        location = multibrand === 'true' ? multibrandLocations : watchLocationsBG;
        break;
      case 'Budapest':
        location = multibrand === 'true' ? [] : watchLocationsBU;
        break;
      case 'Porto Montenegro':
        location = multibrand === 'true' ? [] : watchLocationsPM;
        break;
    }
  }

  // Sort locations
  location.sort();

  const count = location.length;

  return res.status(200).send({
    message: 'Successfully returned product locations',
    count,
    results: location
  });
};

/**
 * @api {get} /product/list/status Get product statuses
 * @apiVersion 1.0.0
 * @apiName getProductStatuses
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String=true} [inStock] Filter products in stock
 * @apiParam (query) {String} [storeId] Filter by Store
 * @apiParam (query) {String} [brand] Filter products by brand
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned product statuses",
   "brand": "Rolex",
   "count": 11,
   "results": [
     "Declined",
     "In progress",
     "New stock",
     "Pre-reserved",
     "Reserved",
     "Staff",
     "Standby",
     "Stock",
     "Vintage",
     "Wishlist",
     "not for sale"
   ]
 }
 * @apiUse InvalidValue
 * @apiUse CredentialsError
 */
 module.exports.getProductStatuses = async (req, res) => {
  const { inStock, storeId, brand } = req.query;

  // Creat 'query' objects
  let queryStock = {};
  let querySoonInStock = {};
  let queryStockBoutiques = {};

  // Check if 'storeId' filter has been sent
  if (storeId) {
    if (!isValidId(storeId)) throw new Error(error.INVALID_VALUE);
    queryStockBoutiques = { 'boutiques.store': ObjectId(storeId) };
    querySoonInStock = { store: storeId };
  }

  // Check if 'brand' filter has been sent
  let productQuery = { brand };

  // Find products, clients and stores by filter data
  const products = await Product.find(productQuery, { _id: 1 }).lean();

  const productFilterApplied = !!Object.keys(productQuery).length;

  // If any of Product is empty(no match), set find by ID to null
  if (productFilterApplied && !products.length) {
    queryStock.product = null;
    querySoonInStock.product = null;
  } else if (productFilterApplied && products.length) {
    // If filter was applied
    queryStock._id = { $in: products.map((p) => p._id) };
    querySoonInStock.product = { $in: products.map((p) => p._id) };
  }

  // Find statuses for whatch in stock and soon in stock
  const [statusesStock, statusesSoonInStock] = await Promise.all([
    Product.aggregate([
      { $match: queryStock },
      { $unwind: '$boutiques' },
      { $match: queryStockBoutiques },
      { $unwind: '$boutiques.serialNumbers' },
      {
        $group: {
          _id: '$boutiques.serialNumbers.status'
        }
      },
      {
        $sort: { 'boutiques.serialNumbers.status': 1 }
      }
    ]),
    SoonInStock.distinct('status', querySoonInStock),
  ]);

  const statusesStockMap = statusesStock.map((a) => a._id);

  let status = statusesStockMap.concat(statusesSoonInStock);

  if (inStock !== 'true') status = statuses;

  // Remove duplicates
  status = status.filter((c, index) => {
    return status.indexOf(c) === index;
  });

  // Sort statuses
  status.sort();

  const count = status.length;

  return res.status(200).send({
    message: 'Successfully returned product statuses',
    brand,
    count,
    results: status
  });
};

/**
 * @api {get} /product/stock-check Print product list
 * @apiVersion 1.0.0
 * @apiName printProductList
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String='Rolex'} brand Filter by Brand Type
 * @apiParam (query) {String} [storeName] Filter by Store Name
 * @apiParam (query) {String='location', 'productLine', 'pgpReference'} [type='location'] Report type
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully printed products list",
   "results": "http://93.186.70.138:3029/rolex.test/product-list-Rolex-Belgrade-2021-02-22d-15-42h.pdf"
 }
 * @apiUse MissingParamsError
 * @apiUse InvalidValue
 */
module.exports.printProductList = async (req, res) => {
  let { store, name } = req.user;
  let { brand, storeName, type = 'location' } = req.query;

  if (!brand) throw new Error(error.MISSING_PARAMETERS);
  if (!brandTypes.includes(brand)) throw new Error(error.INVALID_VALUE);

  storeName = storeName ? storeName : store.name;

  // Creat query object
  // let query = { brand, boutiques: { $elemMatch: { quantity: { $gt: 0 }, storeName } } };
  let query = { brand, boutiques: { $elemMatch: { 'serialNumbers.0': { $exists: true }, storeName } } };

  const pgpReferenceProducts = [];
  if (type === 'pgpReference') {
    const pgpProducts = await Product.find(query, { _id: 1, basicInfo: 1, 'boutiques.$': 1 }).lean();
    for (const pgpProduct of pgpProducts) {
      const [boutique] = pgpProduct.boutiques;
      for (const pgp of boutique.serialNumbers) {
        pgpReferenceProducts.push({
          pgpReference: pgp.pgpReference,
          saleReference: pgpProduct.basicInfo.saleReference,
          dial: pgpProduct.basicInfo.dial,
          bracelet: pgpProduct.basicInfo.bracelet,
          serialNumber: pgp.number,
          eurPrice: boutique.price,
          priceLocal: boutique.priceLocal,
          location: pgp.location,
          materials: pgpProduct.basicInfo.materials,
          jewelryType: pgpProduct.basicInfo.jewelryType,
          size: pgpProduct.basicInfo.size,
          weight: pgpProduct.basicInfo.weight,
          materialDescription: pgpProduct.basicInfo.materialDescription
        });
      }
    }
    pgpReferenceProducts.sort((a, b) => (parseInt(a.pgpReference) > parseInt(b.pgpReference) ? 1 : parseInt(b.pgpReference) > parseInt(a.pgpReference) ? -1 : 0));
  }

  let queryStock = {};
  let queryStockBoutiques = { 'boutiques.storeName': storeName };

  // Find products, clients and stores by filter data
  const products = await Product.find({ brand }, { _id: 1 }).lean();
  queryStock._id = { $in: products.map((p) => p._id) };

  // Find locations for whatch in stock and soon in stock
  const locationsStock = await Product.aggregate([
    { $match: queryStock },
    { $unwind: '$boutiques' },
    { $match: queryStockBoutiques },
    { $unwind: '$boutiques.serialNumbers' },
    {
      $group: {
        _id: '$boutiques.serialNumbers.location'
      }
    },
    {
      $sort: { 'boutiques.serialNumbers.location': 1 }
    }
  ]);

  let printProductLines = {};
  let printCollections = [];
  let printProductByCollection = {};
  let printProductsByCollection = [];

  let locations = locationsStock.map((a) => a._id);
  locations.sort();
  const collections = await Product.distinct('basicInfo.collection', query);

  for (const collection of collections) {
    query['basicInfo.collection'] = collection;
    const productsByCollection = await Product.find(query, { basicInfo: 1, 'boutiques.$': 1 }).sort('basicInfo.saleReference').lean();
    printProductByCollection = {
      name: collection,
      products: productsByCollection,
    };
    printProductsByCollection.push(printProductByCollection);

    const productLines = await Product.distinct('basicInfo.productLine', query);
    for (const productLine of productLines) {
      query['basicInfo.productLine'] = productLine;
      const products = await Product.find(query, { basicInfo: 1, 'boutiques.$': 1 }).sort('basicInfo.saleReference').lean();
      printProductLines = {
        name: `${collection} - ${productLine}`,
        products: products,
      };
      printCollections.push(printProductLines);
    }
    delete query['basicInfo.productLine'];
  }

  let printproductsByLocation = {};
  let printLocations = [];
  delete query['basicInfo.collection'];
  let counter = 0;
  for (const location of locations) {
    query['boutiques.serialNumbers.location'] = location;
    const productsByLocation = await Product.find(query, { basicInfo: 1, 'boutiques.$': 1 }).sort('basicInfo.saleReference').lean();
    for (const productByLocation of productsByLocation) {
      const [productBoutique] = productByLocation.boutiques;
      for (const serialNumberByLocation of productBoutique.serialNumbers) {
        if (serialNumberByLocation.location === location) {
          counter += 1;
        }
      }
    }
    printproductsByLocation = {
      name: location,
      products: productsByLocation
    };
    printLocations.push(printproductsByLocation);
    delete query['boutiques.serialNumbers.location'];
  }
  const printProducts = brand === 'Rolex' || brand === 'Tudor' ? printCollections : printProductsByCollection;
  const url = await stockCheck(printProducts, printLocations, pgpReferenceProducts, storeName, brand, counter, type, name);

  return res.status(200).send({
    message: 'Successfully printed products list',
    results: url
  });
};

/**
 * @api {get} /product/excel Export all products
 * @apiVersion 1.0.0
 * @apiName exportAllProducts
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String='Rolex'} brand Filter by Brand Type
 * @apiParam (query) {String} [collection] Filter by Collection
 * @apiParam (query) {String} [productLine] Filter by Product Line
 * @apiParam (query) {String} [storeId] Filter by Store ID
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully exported products list",
 }
 * @apiUse MissingParamsError
 * @apiUse InvalidValue
 */
module.exports.exportAllProducts = async (req, res) => {
  let { store: userStore } = req.user;
  let { brand, storeId, collection, productLine } = req.query;

  if (!brand) throw new Error(error.MISSING_PARAMETERS);
  if (!brandTypes.includes(brand)) throw new Error(error.INVALID_VALUE);

  storeId = storeId && isValidId(storeId) ? ObjectId(storeId) : userStore._id;
  const store = await Store.findOne({ _id: storeId }).lean();

  // Creat query object
  // let queryStock = { brand, boutiques: { $elemMatch: { quantity: { $gt: 0 }, store: storeId } } };
  let queryStock = { brand, boutiques: { $elemMatch: { 'serialNumbers.0': { $exists: true }, store: storeId } } };

  // Match query for aggregate
  const querySoonInStock = {
    store: storeId
  };

  let collections = await Product.distinct('basicInfo.collection', queryStock);

  if (collection) {
    collections = [collection];
  }

  const workbook = new exceljs.Workbook();
  const worksheet = workbook.addWorksheet('My Sheet');
  worksheet.addRow();
  const font = { name: 'Arial', size: 12 };

  worksheet.columns = [
    {
      header: '#',
      key: 'index',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'PGP',
      key: 'pgp',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'Location',
      key: 'location',
      width: 30,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Status',
      key: 'status',
      width: 30,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Reserved For',
      key: 'reserved',
      width: 30,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'RX Ref',
      key: 'rXRef',
      width: 15,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'RMC',
      key: 'rmc',
      width: 20,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Brac No',
      key: 'bracelet',
      width: 20,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Warranty',
      key: 'warranty',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'S/N',
      key: 'serial',
      width: 15,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'Dial',
      key: 'dial',
      width: 30,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Retail EUR',
      key: 'priceEur',
      width: 15,
      style: { numFmt: '#,##0;[Red]-#,##0', font, alignment: { vertical: 'middle', horizontal: 'right' } }
    },
    {
      header: 'Ex Geneve CHF',
      key: 'exGeneveCHF',
      width: 20,
      style: { numFmt: '#,##0;[Red]-#,##0', font, alignment: { vertical: 'middle', horizontal: 'right' } }
    },
    {
      header: 'Model',
      key: 'model',
      width: 30,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Collection',
      key: 'collection',
      width: 20,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Age',
      key: 'age',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'Store',
      key: 'store',
      width: 15,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    }
  ];

  let count = 0;
  for (const collection of collections) {
    let querySoonInStock2 = {
      'product.basicInfo.collection': collection,
    };
    queryStock['basicInfo.collection'] = collection;
    if (productLine) {
      queryStock['basicInfo.productLine'] = productLine;
      querySoonInStock2['product.basicInfo.productLine'] = productLine;
    }

    // Find products
    const products = await Product.find(queryStock, { 'boutiques.$': 1, basicInfo: 1 }).populate('boutiques.serialNumbers.reservedFor', 'fullName').lean();

    for (const product of products) {
      for (const serialNumber of product.boutiques[0].serialNumbers) {
        const newObject = Object.assign({}, serialNumber);
        newObject.index = count + 1;
        newObject.pgp = newObject.pgpReference;
        newObject.location = newObject.location;
        newObject.status = newObject.status;
        newObject.reserved =  newObject.reservedFor ? newObject.reservedFor.fullName : '';
        newObject.rXRef = product.basicInfo.saleReference;
        newObject.rmc = product.basicInfo.rmc;
        newObject.bracelet = product.basicInfo.bracelet;
        newObject.warranty = newObject.warrantyConfirmed ? 1 : 0;
        newObject.serial = newObject.number;
        newObject.dial = product.basicInfo.dial;
        newObject.priceEur = product.boutiques[0].price;
        newObject.exGeneveCHF = newObject.exGenevaPrice / 100;
        newObject.model = product.basicInfo.productLine;
        newObject.collection = product.basicInfo.collection;
        newObject.age = newObject.stockDate ? Math.ceil(Math.abs(new Date() - new Date(newObject.stockDate)) / (1000 * 30 * 60 * 60 * 24)) : 0;
        newObject.store = store.name;
        worksheet.addRow(newObject, 'i');
        count += 1;
      }
    }

    const soonInStockAggregate = await SoonInStock.aggregate([
      {
        $match: querySoonInStock
      },
      {
        $lookup: {
          from: 'products',
          localField: 'product',
          foreignField: '_id',
          as: 'product'
        }
      },
      { $unwind: '$product' },
      {
        $lookup: {
          from: 'clients',
          localField: 'reservedFor',
          foreignField: '_id',
          as: 'reservedFor'
        }
      },
      {
        $unwind: {
          path: '$reservedFor',
          preserveNullAndEmptyArrays: true
        }
      },
      {
        $match: querySoonInStock2
      },
      {
        $group: {
          _id: '$_id',
          product: { $first: '$$ROOT' }
        }
      },
      {
        $sort: { 'product.product.basicInfo.saleReference': 1 }
      }
    ]);

    for (const soonInStockProduct of soonInStockAggregate) {
      const newObject = Object.assign({}, soonInStockProduct);
      newObject.index = count + 1;
      newObject.pgp = newObject.product.pgpReference;
      newObject.location = newObject.product.location;
      newObject.status = newObject.product.status;
      newObject.reserved = newObject.product.reservedFor ? newObject.product.reservedFor.fullName : '';
      newObject.rXRef = newObject.product.product.basicInfo.saleReference;
      newObject.rmc = newObject.product.rmc;
      newObject.bracelet = newObject.product.product.basicInfo.bracelet;
      newObject.warranty = newObject.product.warrantyConfirmed ? 1 : 0;
      newObject.serial = newObject.product.serialNumber;
      newObject.dial = newObject.product.product.basicInfo.dial;
      newObject.priceEur = newObject.product.product.boutiques[0].price;
      newObject.exGeneveCHF = newObject.product.exGenevaPrice / 100;
      newObject.model = newObject.product.product.basicInfo.productLine;
      newObject.collection = newObject.product.product.basicInfo.collection;
      newObject.age = 0;
      newObject.store = store.name;
      let row = worksheet.addRow(newObject);
      worksheet.getRow(row._number).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DDD9C4' }
      };
      worksheet.getRow(row._number).border = {
        top: { style: 'thin', color: { argb: 'a6a6a6' } },
        left: { style: 'thin', color: { argb: 'a6a6a6' } },
        bottom: { style: 'thin', color: { argb: 'a6a6a6' } },
        right: { style: 'thin', color: { argb: 'a6a6a6' } }
      };

      count += 1;
    }

    worksheet.addRow();
  }

  worksheet.getRow(1).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
  worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
  worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '000' }
  };
  worksheet.getRow(1).border = {
    top: { style: 'thin', color: { argb: 'FFFFFF' } },
    left: { style: 'thin', color: { argb: 'FFFFFF' } },
    bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
    right: { style: 'thin', color: { argb: 'FFFFFF' } }
  };
  worksheet.getRow(1).height = 30;

  const createDir = util.promisify(tmp.dir);
  const tmpDir = await createDir();
  const filePath = `${tmpDir}/${uuidv4()}.xlsx`;

  return workbook.xlsx.writeFile(filePath).then(() => {
    const stream = fs.createReadStream(filePath);

    stream.on('error', () => {
      throw new Error(error.BAD_REQUEST);
    });
    stream.pipe(res);
  });
};

/**
 * @api {get} /product/reservation/excel Export all reserved products
 * @apiVersion 1.0.0
 * @apiName exportReservedProducts
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} [storeId] Filter by Store ID
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully exported reserved products list",
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse InvalidValue
 */
 module.exports.exportReservedProducts = async (req, res) => {
  let { store: userStore } = req.user;
  let { storeId } = req.query;

  storeId = storeId && isValidId(storeId) ? ObjectId(storeId) : userStore._id;
  const store = await Store.findOne({ _id: storeId }).lean();

  if (!store) throw new Error(error.NOT_FOUND);

  // Creat query object
  let queryStock = { 'boutiques.store': storeId, $or: [{ 'boutiques.serialNumbers.status': 'Reserved' }, { 'boutiques.serialNumbers.status': 'Pre-reserved' }] };

  // Match query for aggregate
  const querySoonInStock = { store: storeId, $or: [{ status: 'Reserved' }, { status: 'Pre-reserved' }] };

  const workbook = new exceljs.Workbook();
  const worksheet = workbook.addWorksheet('My Sheet');
  worksheet.addRow();
  const font = { name: 'Arial', size: 12 };

  worksheet.columns = [
    {
      header: '#',
      key: 'index',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'Reservation',
      key: 'reservation',
      width: 15,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
    },
    {
      header: 'PGP Number',
      key: 'pgp',
      width: 20,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'RMC',
      key: 'rmc',
      width: 20,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Dial',
      key: 'dial',
      width: 30,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Bracelet',
      key: 'bracelet',
      width: 20,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Serial',
      key: 'serial',
      width: 15,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'Location',
      key: 'location',
      width: 30,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Date',
      key: 'date',
      width: 15,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Salesperson',
      key: 'salesperson',
      width: 20,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'Reserved for',
      key: 'reservedFor',
      width: 30,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Reservation time',
      key: 'reservationTime',
      width: 18,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
    },
    {
      header: 'Comment',
      key: 'commnet',
      width: 50,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
  ];

  let count = 0;

  // Find products
  const products = await Product.find(queryStock, { 'boutiques.$': 1, basicInfo: 1 }).populate('boutiques.serialNumbers.reservedFor', 'fullName').lean();

  for (const product of products) {
    for (const serialNumber of product.boutiques[0].serialNumbers) {
      if (serialNumber.status === 'Reserved' || serialNumber.status === 'Pre-reserved') {
        const newObject = Object.assign({}, serialNumber);
        const activity = await Activity
          .findOne({ serialNumber: newObject.number, $or: [{ comment: { $regex: '.*Reserved.*' } }, { comment: { $regex: '.*Pre-reserved.*' } }] } )
          .populate('user', 'username')
          .lean();
        newObject.index = count + 1;
        newObject.reservation = serialNumber.status.charAt(0);
        newObject.pgp = newObject.pgpReference;
        newObject.rmc = product.basicInfo.rmc;
        newObject.dial = product.basicInfo.dial;
        newObject.bracelet = product.basicInfo.bracelet;
        newObject.serial = newObject.number;
        newObject.location = newObject.location;
        newObject.date = activity ? activity.createdAt : '';
        newObject.salesperson = activity ? activity.user.username : '';
        newObject.reservedFor = newObject.reservedFor ? newObject.reservedFor.fullName : '';
        newObject.reservationTime = newObject.reservationTime;
        newObject.commnet = newObject.comment;
        worksheet.addRow(newObject, 'i');
        count += 1;
      }
    }
  }

  const soonInStockProducts = await SoonInStock.find(querySoonInStock).populate('reservedFor', 'fullName').populate('product').lean();

  for (const soonInStockProduct of soonInStockProducts) {
    const newObject = Object.assign({}, soonInStockProduct);
    newObject.index = count + 1;
    newObject.reservation = soonInStockProduct.status.charAt(0);
    newObject.pgp = newObject.pgpReference;
    newObject.rmc = newObject.product.basicInfo.rmc;
    newObject.dial = newObject.product.basicInfo.dial;
    newObject.bracelet = newObject.product.basicInfo.bracelet;
    newObject.serial = newObject.serialNumber;
    newObject.location = newObject.location;
    newObject.date = newObject.reservationTime;
    newObject.salesperson = '';
    newObject.reservedFor = newObject.reservedFor ? newObject.reservedFor.fullName : '';
    newObject.commnet = newObject.comment;
    let row = worksheet.addRow(newObject);
    worksheet.getRow(row._number).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'DDD9C4' }
    };
    worksheet.getRow(row._number).border = {
      top: { style: 'thin', color: { argb: 'a6a6a6' } },
      left: { style: 'thin', color: { argb: 'a6a6a6' } },
      bottom: { style: 'thin', color: { argb: 'a6a6a6' } },
      right: { style: 'thin', color: { argb: 'a6a6a6' } }
    };

    count += 1;
  }

  worksheet.addRow();

  worksheet.getRow(1).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
  worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
  worksheet.getRow(1).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: '000' }
  };
  worksheet.getRow(1).border = {
    top: { style: 'thin', color: { argb: 'FFFFFF' } },
    left: { style: 'thin', color: { argb: 'FFFFFF' } },
    bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
    right: { style: 'thin', color: { argb: 'FFFFFF' } }
  };
  worksheet.getRow(1).height = 30;

  const createDir = util.promisify(tmp.dir);
  const tmpDir = await createDir();
  const filePath = `${tmpDir}/${uuidv4()}.xlsx`;

  return workbook.xlsx.writeFile(filePath).then(() => {
    const stream = fs.createReadStream(filePath);

    stream.on('error', () => {
      throw new Error(error.BAD_REQUEST);
    });
    stream.pipe(res);
  });
};

/**
 * @api {get} /product/reservation/excel/soonInStock Export reserved SoonInStocks
 * @apiVersion 1.0.0
 * @apiName exportReservedSoonInStocks
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} [storeId] Filter by Store ID
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully exported reserved SoonInStocks",
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse InvalidValue
 */
 module.exports.exportReservedSoonInStocks = async (req, res) => {
  let { store: userStore } = req.user;
  let { storeId } = req.query;

  storeId = storeId && isValidId(storeId) ? ObjectId(storeId) : userStore._id;
  const store = await Store.findOne({ _id: storeId }).lean();

  if (!store) throw new Error(error.NOT_FOUND);

  // Create excel workbook
  const workbook = new exceljs.Workbook();

  // Create sheet
  const worksheet = workbook.addWorksheet(`Reserved sooninstocks ${store.name}`);
  worksheet.addRow();

  // Set font
  const font = { name: 'Arial', size: 12 };

  // Set columns
  worksheet.columns = [
    {
      header: 'PGP',
      key: 'pgpReference',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'Location',
      key: 'location',
      width: 20,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Status',
      key: 'status',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Reserved For - Client ID',
      key: 'reservedFor',
      width: 26,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Reservation Time',
      key: 'reservationTime',
      width: 16,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'Comment',
      key: 'comment',
      width: 30,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'Previous Serial',
      key: 'serialNumber',
      width: 17,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'soldToParty',
      key: 'soldToParty',
      width: 11,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'shipToParty',
      key: 'shipToParty',
      width: 11,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'billToParty',
      key: 'billToParty',
      width: 11,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'packingNumber',
      key: 'packingNumber',
      width: 14,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'invoiceDate',
      key: 'invoiceDate',
      width: 12,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'invoiceNumber',
      key: 'invoiceNumber',
      width: 16,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
    },
    {
      header: 'Serial',
      key: 'serialNumber2',
      width: 17,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'boxCode',
      key: 'boxCode',
      width: 12,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'exGenevaPrice',
      key: 'exGenevaPrice',
      width: 14,
      style: { font, alignment: { vertical: 'middle', horizontal: 'right' } }
    },
    {
      header: 'sectorId',
      key: 'sectorId',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'sector',
      key: 'sector',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'boitId',
      key: 'boitId',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'boitRef',
      key: 'boitRef',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'center' } }
    },
    {
      header: 'caspId',
      key: 'caspId',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'cadranDesc',
      key: 'cadranDesc',
      width: 20,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'ldisId',
      key: 'ldisId',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'ldisq',
      key: 'ldisq',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'brspId',
      key: 'brspId',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'brspRef',
      key: 'brspRef',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'prixSuisse',
      key: 'prixSuisse',
      width: 12,
      style: { font, alignment: { vertical: 'middle', horizontal: 'right' } }
    },
    {
      header: 'RMC',
      key: 'rmc',
      width: 18,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'umIntern',
      key: 'umIntern',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'umExtern',
      key: 'umExtern',
      width: 11,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'rfid',
      key: 'rfid',
      width: 10,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    },
    {
      header: 'packingType',
      key: 'packingType',
      width: 13,
      style: { font, alignment: { vertical: 'middle', horizontal: 'left' } }
    }
  ];

  const soonInStockProducts = await SoonInStock.find({ store: storeId, $or: [{ status: 'Reserved' }, { status: 'Pre-reserved' }] }).lean();

  for (const soon of soonInStockProducts) {
    const newObject = Object.assign({}, soon);

    newObject.serialNumber2 = soon.serialNumber;
    newObject.reservedFor = soon.reservedFor ? soon.reservedFor.toString() : soon.reservedFor;
    newObject.reservationTime = soon.reservationTime ? new Date(soon.reservationTime).toISOString().slice(0, 10) : soon.reservationTime;
    newObject.invoiceDate = soon.invoiceDate ? new Date(soon.invoiceDate).toISOString().slice(0, 10).split('-').join('') : soon.reservationTime;

    let row = worksheet.addRow(newObject);
    // let row = worksheet.addRow(soon);
  }

  for (let i = 1; i < 8; i++) {
    worksheet.getColumn(i).eachCell((cell, rowNumber) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'DACBAB' }
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFFFFF' } },          // WHITE
        left: { style: 'thin', color: { argb: 'FFFFFF' } },
        bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
        right: { style: 'thin', color: { argb: 'FFFFFF' } },
      };
    });
  }

  for (let i = 8; i < 33; i++) {
    worksheet.getColumn(i).eachCell((cell, rowNumber) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'E9F1EC' }
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFFFFF' } },
        left: { style: 'thin', color: { argb: 'FFFFFF' } },
        bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
        right: { style: 'thin', color: { argb: 'FFFFFF' } },
      };
    });
  }

  worksheet.getColumn(14).eachCell((cell, rowNumber) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'c1d8c9' }
    };
  });

  worksheet.getColumn(7).eachCell((cell, rowNumber) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'cab588' }
    };
  });

  for(let i = 1; i < 33; i++) {
    worksheet.getRow(1).getCell(i).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '006039' }
    };
    worksheet.getRow(1).getCell(i).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
    worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
    worksheet.getRow(1).getCell(i).border = {
      top: { style: 'thin', color: { argb: 'FFFFFF' } },
      left: { style: 'thin', color: { argb: 'FFFFFF' } },
      bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
      right: { style: 'thin', color: { argb: 'FFFFFF' } }
    };
  }

  worksheet.getRow(1).height = 26;

  const createDir = util.promisify(tmp.dir);
  const tmpDir = await createDir();
  const filePath = `${tmpDir}/${uuidv4()}.xlsx`;

  return workbook.xlsx.writeFile(filePath).then(() => {
    const stream = fs.createReadStream(filePath);

    stream.on('error', () => {
      throw new Error(error.BAD_REQUEST);
    });
    stream.pipe(res);
  });
};

/**
 * @api {get} /product/serial-number Get product by serial number
 * @apiVersion 1.0.0
 * @apiName getProductBySerialNumber
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} serialNumber Serial number
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned product by serial number",
   "serialNumber": {
    "warrantyConfirmed": true,
     "modified": false,
     "_id": "603cac39646fec5688d3f418",
     "number": "346887F8",
     "pgpReference": "6797",
     "status": "Vintage",
     "stockDate": "2020-07-10T04:00:00.000Z",
     "location": "Room B Big safe Right",
     "comment": "",
     "card": "322",
     "soldToParty": "322",
     "shipToParty": "322",
     "billToParty": "32200",
     "packingNumber": "110366937",
     "invoiceDate": "1970-01-01T05:36:40.710Z",
     "invoiceNumber": "7239560",
     "boxCode": "OYSTER M",
     "exGenevaPrice": 395000,
     "sectorId": "10",
     "sector": "RO",
     "boitId": "41343",
     "boitRef": "114060",
     "caspId": "42290",
     "cadranDesc": "BLACK INDEX W",
     "ldisId": "0",
     "ldisq": "",
     "brspId": "41315",
     "brspRef": "97200",
     "prixSuisse": 750000,
     "umIntern": "152966454",
     "umExtern": "193777-001",
     "rfid": "",
     "packingType": "ECU",
     "reservedFor": {
       "_id": "6167eabcb5bf4c1bb016200e",
       "birthDate": {
         "day": 21,
         "month": "May",
         "year": 1991
       },
       "passport": {},
       "dashboard": {
         "directTurnover": {
           "net": 0,
           "taxes": 0,
           "referral": 0
         },
         "rolexPurchasedProducts": 7,
         "rolexOffersReceived": 0,
         "multibrandPurchasedProducts": 0,
         "multibrandOffersReceived": 0,
         "afterSales": 0,
         "selfUsedPurchasedProducts": 0,
         "giftedPurchasedProducts": 0,
         "uniqueReferrals": 0,
         "productsGifted": 0,
         "brandedGiftsReceived": 0
       },
       "pgpHistory": [
         "C-00016879"
       ],
       "passports": [],
       "clientClass": "vip",
       "lockClientClass": false,
       "lockCountryOfRef": false,
       "isDiplomat": null,
       "blacklisted": "clear",
       "blacklistComment": "",
       "locked": false,
       "documents": [
         "61a9cdef4bea23cca5458b5d"
       ],
       "numberOfWishlists": 4,
       "notes": [],
       "badges": [],
       "offers": [],
       "numberOfOffers": 0,
       "purchases": [],
       "profileStrength": 5,
       "merged": false,
       "active": true,
       "loyaltyPoints": 524,
       "clientType": "individual",
       "pgpReference": "C-00016879",
       "fullName": "Nikola Ljubiša Đorđević",
       "firstName": "Nikola",
       "middleName": "Ljubiša",
       "lastName": "Đorđević",
       "companyName": "",
       "nameSearch": "nikola ljubiša đorđević; nikola ljubisa djordjevic; ljubiša nikola đorđević; ljubisa nikola djordjevic; đorđević nikola ljubiša; djordjevic nikola ljubisa; nikola đorđević ljubiša; nikola djordjevic ljubisa; ljubiša đorđević nikola; ljubisa djordjevic nikola; đorđević ljubiša nikola; djordjevic ljubisa nikola;",
       "companyType": "",
       "photo": "1591977301736-1634200034775.jpeg",
       "title": "Dr.",
       "addresses": [
         {
           "main": true,
           "verified": null,
           "street": "Zvezde Danice, Surčin, Serbia",
           "country": "Serbia",
           "countryCode": "RS",
           "city": "Surčin",
           "postalCode": "11271",
           "streetShort": "Zvezde Danice 6a/9"
         }
       ],
       "phoneMain": {
         "contactType": "Mobile",
         "country": "",
         "prefix": "+381",
         "number": "66008576",
         "fullNumber": "38166008576"
       },
       "phones": [
         {
           "main": true,
           "verified": false,
           "contactType": "Mobile",
           "country": "",
           "prefix": "+381",
           "number": "66008576",
           "fullNumber": "38166008576"
         }
       ],
       "emails": [
         {
           "main": true,
           "verified": false,
           "contactType": "Work",
           "address": "nikola.djordjevic@30hills.com"
         }
       ],
       "geoInfo": "local",
       "countryOfRef": "Serbia",
       "countryOfRefComment": "",
       "boutiqueOfRef": "Belgrade",
       "alias": "Nik",
       "socialNetworks": [],
       "introducedBy": null,
       "introducedByFreeForm": "",
       "workInformation": [],
       "family": [],
       "network": [],
       "clientTags": [],
       "productTags": [],
       "identityRegNumber": "210599175101",
       "taxNumber": "",
       "companyWebsite": "",
       "employees": [],
       "diplomaticIdCardNumber": "",
       "diplomaticCountry": "",
       "customerSince": "2021-10-14T08:30:50.519Z",
       "quickNote": "Back End Developer at 30Hills",
       "responsiblePerson": "6156e14ffa2cb00a0ece3403",
       "responsiblePersonName": "test admin",
       "createdBy": "test admin",
       "lastEventDate": "2021-12-08T08:16:46.518Z",
       "gifts": [],
       "wishlists": [],
       "createdAt": "2021-10-14T08:30:52.855Z",
       "updatedAt": "2021-12-08T08:18:24.038Z",
       "__v": 0,
       "gdpr": "61a9cdef4bea23cca5458b5d",
       "specialClientOccupation": "athlete"
     },
     "reservationTime": "2021-12-11T08:26:48.624Z",
   },
   "results": {
     "_id": "5f92b07ea79cbc2117179922",
     "basicInfo": {
       "photos": [
         "http://content.rolex.com/dam/2020/upright-bba-with-shadow/m116769tbrj-0002.png?impolicy=v6-upright&imwidth=420",
       ],
       "rmc": "M116769TBRJ-0002",
       "collection": "PROFESSIONAL",
       "productLine": "GMT-MASTER II",
       "saleReference": "116769TBRJ",
       "materialDescription": "PAVED W-74779BRJ",
       "dial": "PAVED W",
       "bracelet": "74779BRJ",
       "box": "EN DD EMERAUDE 60",
       "exGeneveCHF": 1565800,
       "diameter": 40,
     },
     "brand": "Rolex",
     "status": "new",
     "active": true,
     "boutiques": [
       {
         "quantity": 1,
         "_id": "5f92b07ea79cbc2117179923",
         "store": "5f92b07ea79cbc211717991e",
         "storeName": "Belgrade",
         "serialNumbers": [
           {
             "warrantyConfirmed": true,
             "modified": false,
             "_id": "603cac39646fec5688d3f418",
             "number": "346887F8",
             "pgpReference": "6797",
             "status": "Vintage",
             "stockDate": "2020-07-10T04:00:00.000Z",
             "location": "Room B Big safe Right",
             "comment": "",
             "card": "322",
             "soldToParty": "322",
             "shipToParty": "322",
             "billToParty": "32200",
             "packingNumber": "110366937",
             "invoiceDate": "1970-01-01T05:36:40.710Z",
             "invoiceNumber": "7239560",
             "boxCode": "OYSTER M",
             "exGenevaPrice": 395000,
             "sectorId": "10",
             "sector": "RO",
             "boitId": "41343",
             "boitRef": "114060",
             "caspId": "42290",
             "cadranDesc": "BLACK INDEX W",
             "ldisId": "0",
             "ldisq": "",
             "brspId": "41315",
             "brspRef": "97200",
             "prixSuisse": 750000,
             "umIntern": "152966454",
             "umExtern": "193777-001",
             "rfid": "",
             "packingType": "ECU"
           }
         ]
       }
     ],
     "createdAt": "2020-10-23T10:29:18.533Z",
     "updatedAt": "2020-10-23T10:29:18.533Z",
     "__v": 0
   }
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 */
 module.exports.getProductBySerialNumber = async (req, res) => {
  const { serialNumber } = req.query;

  if (!serialNumber) throw new Error(error.MISSING_PARAMETERS);

  // Find product
  const product = await Product.findOne({ 'boutiques.serialNumbers.number': serialNumber }, { 'boutiques.$': 1, basicInfo: 1, brand: 1, status: 1, active: 1 }).populate('boutiques.serialNumbers.reservedFor').lean();

  // Check if product is found
  if (!product) throw new Error(error.NOT_FOUND);

  // Find the right store
  const [boutique] = product.boutiques;
  const number = boutique.serialNumbers.find((serial) => serial.number === serialNumber);

  return res.status(200).send({
    message: 'Successfully returned product by serial number',
    serialNumber: number,
    results: product,
  });
};

 /**
 * @api {patch} /product/:productId/serialnumber/change-rmc Change RMC of a watch
 * @apiVersion 1.0.0
 * @apiName changeRmc
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - User with permission
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (body) {String} serialNumber Serial Number
 * @apiParam (body) {String} newRMC New RMC of a watch
 * @apiParam (body) {String} [comment] Comment
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
  {
    "message": "Successfully changed RMC of watch with serial number: 0680X156 from M126333-0002 to M126333-0014",
    "results": {
      "_id": "603caa9a76077912220babf8",
      "basicInfo": {
          "photos": [
              "http://192.168.12.50:9000/minio/download/rolex/M126333-0014.jpg?token="
          ],
          "materials": [],
          "rmc": "M126333-0014",
          "collection": "OYSTER",
          "productLine": "DATEJUST",
          "saleReference": "126333",
          "materialDescription": "BRIGHT BLACK INDEX Y-62613",
          "dial": "BRIGHT BLACK INDEX Y",
          "bracelet": "62613",
          "box": "OYSTER M",
          "exGeneveCHF": 6868,
          "diameter": "41",
          "stones": [],
          "diamonds": [],
          "braceletType": "Jubilee",
          "caseMaterial": "BRIGHT BLACK INDEX Y"
        },
        "quotaRegime": false,
        "status": "new",
        "brand": "Rolex",
        "boutiques": [
          {
            "quantity": 0,
            "_id": "607b3b48d655784ccdc8cf8f",
            "store": "5f3a4225ffe375404f72fb08",
            "storeName": "Porto Montenegro",
            "price": 12600,
            "priceLocal": 12600,
            "VATpercent": 21,
            "priceHistory": [
              {
                "_id": "607b3b48d655784ccdc8cf93",
                "date": "2021-04-07T00:00:00.000Z",
                "price": 12600,
                "VAT": 21,
                "priceLocal": 12600
              }
            ],
            "serialNumbers": []
        },
        {
          "quantity": 1,
          "_id": "607b3b48d655784ccdc8cf94",
          "store": "5f3a4225ffe375404f72fb07",
          "storeName": "Budapest",
          "price": 13200,
          "priceLocal": 4950000,
          "VATpercent": 27,
          "priceHistory": [
            {
              "_id": "607b3b48d655784ccdc8cf98",
              "date": "2021-04-07T00:00:00.000Z",
              "price": 13200,
              "VAT": 27,
              "priceLocal": 4950000
            }
          ],
          "serialNumbers": [
            {
              "warrantyConfirmed": true,
              "modified": false,
              "_id": "603cac35646fec5688d3eebc",
              "number": "0680X156",
              "pgpReference": "7524",
              "status": "Reserved",
              "stockDate": "2021-01-19T05:00:00.000Z",
              "location": "Room B Small safe",
              "comment": "Changing dials for Sasha // reserved by AP for Dr. Guba Áron",
              "card": "322",
              "soldToParty": "322",
              "shipToParty": "322",
              "billToParty": "32200",
              "packingNumber": "110387286",
              "invoiceDate": "1970-01-01T05:36:50.119Z",
              "invoiceNumber": "7251870",
              "boxCode": "OYSTER M",
              "exGenevaPrice": 663500,
              "sectorId": "10",
              "sector": "RO",
              "boitId": "50642",
              "boitRef": "126333",
              "caspId": "50646",
              "cadranDesc": "SILVER INDEX Y",
              "ldisId": "0",
              "ldisq": "",
              "brspId": "50651",
              "brspRef": "62613",
              "prixSuisse": 1260000,
              "umIntern": "153204090",
              "umExtern": "205117-002",
              "rfid": "",
              "packingType": "ECU",
              "reservationTime": "2021-04-18T22:00:00.000Z",
              "reservedFor": "607da8df325ea3a85171e897"
            }
          ]
        },
        {
          "quantity": 0,
          "_id": "607b3b48d655784ccdc8cf99",
          "store": "5f3a4225ffe375404f72fb06",
          "storeName": "Belgrade",
          "price": 12450,
          "priceLocal": 1494000,
          "VATpercent": 20,
          "priceHistory": [
            {
              "_id": "607b3b48d655784ccdc8cf9d",
              "date": "2021-04-07T00:00:00.000Z",
              "price": 12450,
              "VAT": 20,
              "priceLocal": 1494000
            }
          ],
          "serialNumbers": []
        }
      ],
      "__v": 0,
      "createdAt": "2021-03-01T08:49:40.895Z",
      "updatedAt": "2021-04-20T17:25:50.368Z",
      "wishlist": "603cafe60d99a81393c7c590",
      "active": true
    }
  }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse NotAcceptable
 * @apiUse CredentialsError
 */
  module.exports.changeRmc = async (req, res) => {
    const { productId } = req.params;
    const { serialNumber, newRMC } = req.body;
    let { comment: sentComment } = req.body;
    const { _id: userId } = req.user;

    // Check if all required parameters have been sent
    if (!serialNumber || !newRMC) throw new Error(error.MISSING_PARAMETERS);

    // Find product that contains sent 'serialNumber'
    const currentProduct = await Product.findOne({
      _id: productId,
      'boutiques.serialNumbers.number': serialNumber
    }).lean();

    // Check if product was found
    if (!currentProduct) throw new Error(error.NOT_FOUND);

    // Check if attempts to change to RMC it already belongs to
    if (currentProduct.basicInfo.rmc === newRMC) throw new Error(error.NOT_ACCEPTABLE);

    // Get related watch object in product 'boutiques' array
    let watch = {};
    let currentStore = '';

    for (const boutique of currentProduct.boutiques) {
      const watches = boutique.serialNumbers.filter(sn => sn.number === serialNumber);
      if (watches.length > 0) {
        [watch] = watches;
        currentStore = boutique.storeName;
      }
    }

    // 1. Update product -> push watch object to new RMC
    const addWatchToProduct = await Product.findOneAndUpdate(
      { 'basicInfo.rmc': newRMC, boutiques: { $elemMatch: { storeName: currentStore } } },
      {
        $addToSet:
          { 'boutiques.$.serialNumbers': watch },
        $inc: { 'boutiques.$.quantity': 1 }
      },
      { new: true }
    ).lean();

    // Create new activity
    let manuallyAdded = true;

    if (!sentComment) {
      sentComment = '';
      manuallyAdded = false;
    }

    const comment = `Changed RMC from '${currentProduct.basicInfo.rmc}' to '${newRMC}'. ${sentComment}`;
    const newActivity = createActivity('Product', userId, null, comment, null, addWatchToProduct._id, serialNumber, null, new Date(), manuallyAdded);

    // Execute
    await Promise.all([
      // 2. Update product -> remove watch object from current location and
      Product.updateOne(
        { _id: productId, boutiques: { $elemMatch: { storeName: currentStore } } },
        {
          $pull: { 'boutiques.$.serialNumbers': { number: serialNumber } },
          $inc: { 'boutiques.$.quantity': -1 }
        }
      ),
      // 3. Update Activities -> switch product ID
      Activity.updateMany(
        {
          serialNumber: watch.number,
          type: 'Product'
        },
        { $set: { product: addWatchToProduct._id } }
      ),
      newActivity.save()
    ]);

    return res.status(200).send({
      message: `Successfully changed watch RMC`,
      results: addWatchToProduct
    });
  };

 /**
 * @api {get} /product/list/rmc Get RMCs by sale reference
 * @apiVersion 1.0.0
 * @apiName getRmcsByReference
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} saleReference Sale reference
 * @apiParam (query) {String='Belgrade', 'Budapest', 'Porto Montenegro'} storeId Store ID
 * @apiParam (query) {String} brand Brand
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
  {
    "message": "Successfully returned list of RMCs",
    "results": [
      {
        "rmc": "M126234-0011",
        "title": "M126234-0011 --- 126234 --- BRIGHT BLUE JUB 10BR W --- 62800 --- EUR 9950"
      },
      {
        "rmc": "M126234-0012",
        "title": "M126234-0012 --- 126234 --- BRIGHT BLUE JUB 10BR W --- 72800 --- EUR 9750"
      },
      {
        "rmc": "M126234-0013",
        "title": "M126234-0013 --- 126234 --- SILVER INDEX W --- 62800 --- EUR 8050"
      },
      ...
    ]
  }
 * @apiUse MissingParamsError
 * @apiUse InvalidValue
 * @apiUse CredentialsError
 */
module.exports.getRmcsByReference = async (req, res) => {
  const { saleReference, storeId, brand } = req.query;

  // Check if required data has been sent
  if (!saleReference || !storeId || !brand) throw new Error(error.MISSING_PARAMETERS);

  // Check if 'storeId' is correct
  const storeIds = await Store.distinct('_id');
  if (!storeIds.map(el => el.toString()).includes(storeId)) throw new Error(error.INVALID_VALUE);

  // Check if 'brand' is correct
  if (!brandTypes.includes(brand)) throw new Error(error.INVALID_VALUE);

  // Fetch all RMCs related to the sent 'saleReference'
  const products = await Product.find(
    { brand, 'basicInfo.saleReference': saleReference }
  ).lean();

  const results = [];

  for (const product of products) {
    const [boutique] = product.boutiques.filter(boutique => boutique.store.toString() === storeId.toString());
    results.push({ rmc: product.basicInfo.rmc, title: `${product.basicInfo.rmc} --- ${product.basicInfo.dial} --- ${product.basicInfo.bracelet} --- EUR ${boutique.price}` });
  }

  return res.status(200).send({
    message: 'Successfully returned list of RMCs',
    results
  });
};

/**
 * @api {get} /product/pgp-reference Get watch by pgp reference
 * @apiVersion 1.0.0
 * @apiName getPgpReference
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String} pgpReference PGP reference / Serial Number
 * @apiParam (query) {String='Sale'} type Filter products by type
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully returned product by pgp reference / serial number",
   "count": 1,
   "results": [
     {
       "_id": "603caa9776077912220b18a2",
       "boutique": "Belgrade",
       "brand": "Rolex",
       "serialNumber": {
         "warrantyConfirmed": false,
         "modified": false,
         "_id": "60873ed5a40a78584756763e",
         "number": "346887F8",
         "pgpReference": "10673",
         "status": "Vintage",
         "stockDate": "2020-07-09T22:00:00.000Z",
         "location": "ACC safe",
         "comment": "",
         "origin": "Geneva",
         "card": "322",
         "transferDate": null,
         "soldToParty": "322",
         "shipToParty": "322",
         "billToParty": "32200",
         "packingNumber": "110366937",
         "invoiceDate": "1970-01-01T05:36:40.710Z",
         "invoiceNumber": "7239560",
         "boxCode": "OYSTER M",
         "exGenevaPrice": 395000,
         "sectorId": "10",
         "sector": "RO",
         "boitId": "41343",
         "boitRef": "114060",
         "caspId": "42290",
         "cadranDesc": "BLACK INDEX W",
         "ldisId": "0",
         "ldisq": "",
         "brspId": "41315",
         "brspRef": "97200",
         "prixSuisse": 750000,
         "umIntern": "152966454",
         "umExtern": "193777-001",
         "rfid": "",
         "packingType": "ECU"
       },
       "basicInfo": {
         "photos": [
           "M114060-0002.jpg"
         ],
         "materials": [],
         "rmc": "M114060-0002",
         "collection": "VINTAGE",
         "productLine": "SUBMARINER",
         "saleReference": "114060",
         "materialDescription": "BLACK INDEX W-97200",
         "dial": "BLACK INDEX W",
         "bracelet": "97200",
         "box": "OYSTER M",
         "exGeneveCHF": 3950,
         "diameter": "40",
         "stones": [],
         "diamonds": [],
         "braceletType": "Oyster",
         "caseMaterial": "Black"
       }
     }
   ]
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 */
 module.exports.getPgpReference = async (req, res) => {
  const { pgpReference, type } = req.query;

  if (!pgpReference) throw new Error(error.MISSING_PARAMETERS);

  // Find product, soonInStockProducts and salesReport
  const [products, soonInStockProducts, salesReports] = await Promise.all([
    Product.find({ $or: [{ 'boutiques.serialNumbers.number': new RegExp(pgpReference, 'i') }, { 'boutiques.serialNumbers.pgpReference': new RegExp(pgpReference, 'i') }] }, { boutiques: 1, basicInfo: 1, brand: 1 }).lean(),
    SoonInStock.find({ $or: [{ serialNumber: new RegExp(pgpReference, 'i') }, { pgpReference: new RegExp(pgpReference, 'i') }] })
      .populate('product', 'basicInfo brand')
      .populate('store', 'name')
      .lean(),
    Report.find({ $or: [{ 'serialNumber.number': new RegExp(pgpReference, 'i') }, { 'serialNumber.pgpReference': new RegExp(pgpReference, 'i') }] })
      .populate('product', 'basicInfo brand')
      .populate('store', 'name')
      .lean()
  ]);

  // Check if product is found
  if (!products.length && !soonInStockProducts.length && !salesReports.length) throw new Error(error.NOT_FOUND);

  let results = [];
  if (products) {
   for (const product of products) {
      // Find the right store
      const boutique = product.boutiques.find((boutique) => boutique.serialNumbers.find((serial) => serial.number.includes(pgpReference) || serial.pgpReference.includes(pgpReference)));

      // Find serial number
      const number = boutique.serialNumbers.find((serial) => serial.number.includes(pgpReference) || serial.pgpReference.includes(pgpReference));

      results.push({
        type: 'Stock',
        _id: product._id,
        boutique: boutique.storeName,
        brand: product.brand,
        serialNumber: number,
        basicInfo: product.basicInfo,
      });
    }
  }

  if (soonInStockProducts) {
    for (const soonInStock of soonInStockProducts) {
      results.push({
        type: 'Soon In Stock',
        _id: soonInStock.product._id,
        boutique: soonInStock.store.name,
        brand: soonInStock.product.brand,
        serialNumber: {
          warrantyConfirmed: soonInStock.warrantyConfirmed,
          number: soonInStock.serialNumber,
          pgpReference: soonInStock.pgpReference,
          status: soonInStock.status,
          location: soonInStock.location,
          comment: soonInStock.comment,
          card: soonInStock.card,
          soldToParty: soonInStock.soldToParty,
          shipToParty: soonInStock.shipToParty,
          billToParty: soonInStock.billToParty,
          packingNumber: soonInStock.packingNumber,
          invoiceDate: soonInStock.invoiceDate,
          invoiceNumber: soonInStock.invoiceNumber,
          boxCode: soonInStock.boxCode,
          exGenevaPrice: soonInStock.exGenevaPrice,
          sectorId: soonInStock.sectorId,
          sector: soonInStock.sector,
          boitId: soonInStock.boitId,
          boitRef: soonInStock.boitRef,
          caspId: soonInStock.caspId,
          cadranDesc: soonInStock.cadranDesc,
          ldisId: soonInStock.ldisId,
          ldisq: soonInStock.ldisq,
          brspId: soonInStock.brspId,
          brspRef: soonInStock.brspRef,
          prixSuisse: soonInStock.prixSuisse,
          umIntern: soonInStock.umIntern,
          umExtern: soonInStock.umExtern,
          rfid: soonInStock.rfid,
          packingType: soonInStock.packingType,
          reservedFor: soonInStock.reservedFor
        },
        basicInfo: soonInStock.product.basicInfo
      });
    }
  }

  // Empty the array if only sales reports are required
  if (type === 'Sale') results = [];

  if (salesReports) {
    for (const salesReport of salesReports) {
      let owner, referral, purchaser;
      if (type === 'Sale') {
        [owner, referral, purchaser] = await Promise.all([
          Client.findOne({ _id: salesReport.owner.clientId }, { blacklisted: 1, blacklistComment: 1 }).lean(),
          Client.findOne({ _id: salesReport.referral.clientId }, { blacklisted: 1, blacklistComment: 1 }).lean(),
          Client.findOne({ _id: salesReport.purchaser.clientId }, { blacklisted: 1, blacklistComment: 1 }).lean()
        ]);
      }

      results.push({
        type: 'Sale',
        _id: salesReport.product ? salesReport.product._id : '',
        report: salesReport._id,
        boutique: salesReport.store.name,
        brand: salesReport.brand,
        serialNumber: salesReport.serialNumber,
        basicInfo: salesReport.product ? salesReport.product.basicInfo : '',
        owner: salesReport.owner,
        ownerBlacklisted: owner ? owner.blacklisted : 'clear',
        ownerBlacklistedComment: owner ? owner.blacklistComment : '',
        referral: salesReport.referral,
        referralBlacklisted: referral ? referral.blacklisted : 'clear',
        referralBlacklistedComment: referral ? referral.blacklistComment : '',
        purchaser: salesReport.purchaser,
        purchaserBlacklisted: purchaser ? purchaser.blacklisted : 'clear',
        purchaserBlacklistedComment: purchaser ? purchaser.blacklistComment : '',
        salesDate: moment(salesReport.salesDate).clone().format('DD.MM.YYYY'),
        isResold: salesReport.isResold ? salesReport.isResold : false,
        resoldComment: salesReport.resoldComment ? salesReport.resoldComment : '',
        verified: salesReport.verified ? salesReport.verified : false
      });
    }
  }

  return res.status(200).send({
    message: 'Successfully returned product by pgp reference / serial number',
    count: results.length,
    results
  });
};

/**
 * @api {patch} /product/:productId/sold Declare product as sold
 * @apiVersion 1.0.0
 * @apiName declareAsSold
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (params) {String} productId Product ID
 * @apiParam (query) {String} serialNumber Serial number
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully declared product as sold",
   "results": {
     "serialNumber": {
       "modified": false,
       "warrantyConfirmed": true,
       "number": "34F5U993",
       "pgpReference": "10727",
       "location": "SAV safe",
       "status": "New stock",
       "comment": "",
       "stockDate": "2021-05-07T00:00:00.000Z",
       "origin": "Geneva",
       "card": "325",
       "soldToParty": "325",
       "shipToParty": "325",
       "billToParty": "32500",
       "packingNumber": "110400337",
       "invoiceDate": "1970-01-01T05:36:50.503Z",
       "invoiceNumber": "7259731",
       "boxCode": "OYSTER M",
       "exGenevaPrice": 501500,
       "sectorId": "10",
       "sector": "RO",
       "boitId": "51198",
       "boitRef": "126710BLNR",
       "caspId": "51437",
       "cadranDesc": "BLACK WHITE PRINT IND W",
       "ldisId": "0",
       "ldisq": "",
       "brspId": "51438",
       "brspRef": "69200",
       "prixSuisse": 920000,
       "umIntern": "153363891",
       "umExtern": "212248-002",
       "rfid": "",
       "packingType": "ECU"
     },
     "payment": {
       "card": 0,
       "wireTransfer": 0,
       "cash": 0
     },
     "owner": {
       "anonymousClient": false,
       "phones": [],
       "emails": [],
       "addresses": []
     },
     "referral": {
       "anonymousClient": false,
       "phones": [],
       "emails": [],
       "addresses": []
     },
     "purchaser": {
       "anonymousClient": false,
       "phones": [],
       "emails": [],
       "addresses": []
     },
     "restricted": false,
     "language": "sr",
     "isDiplomaticSale": false,
     "warranty": false,
     "currency": "rsd",
     "taxExemption": false,
     "documents": [],
     "boxes": [],
     "gifts": [],
     "services": [],
     "manuallyCreated": true,
     "_id": "60d1acc138e808273b346d6f",
     "ordinal": 6328,
     "status": "requires data",
     "store": "5f3a4225ffe375404f72fb06",
     "createdBy": "603ca9d122c0d111a0b4242c",
     "responsiblePerson": "603ca9d122c0d111a0b4242c",
     "product": "603caa9a76077912220bae45",
     "brand": "Rolex",
     "salesType": "Sale",
     "retailPrice": 9100,
     "discount": 0,
     "finalPrice": 0,
     "vat": 0,
     "netPrice": 0,
     "invoicedLocalCurrency": 0,
     "salesDate": "2021-06-22T09:26:25.950Z",
     "createdAt": "2021-06-22T09:26:26.033Z",
     "updatedAt": "2021-06-22T09:26:26.033Z",
     "__v": 0
   }
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 */
module.exports.declareAsSold = async (req, res) => {
  const { _id: createdBy } = req.user;
  const { productId } = req.params;
  const { serialNumber } = req.query;

  // Check if all required parameters have been sent
  if (!serialNumber) throw new Error(error.MISSING_PARAMETERS);

  // Find product
  const product = await Product.findOne({ _id: productId, 'boutiques.serialNumbers.number': serialNumber }, { 'boutiques.$': 1, basicInfo: 1, brand: 1, status: 1, active: 1 }).lean();

  // Check if product is found
  if (!product) throw new Error(error.NOT_FOUND);

  // Find the right store
  const [boutique] = product.boutiques;

  // Find serial number object
  const serialNumberObj = boutique.serialNumbers.find((serial) => serial.number === serialNumber);

  // Set report ordinal number
  let reportOrdinal = 1;
  const [[lastReport], reservedForClient] = await Promise.all([
    Report.find({ store: boutique.store }).sort({ ordinal: -1 }).lean(),
    Client.findOne({ _id: serialNumberObj.reservedFor }).lean(),
  ]);
  if (lastReport) {
    reportOrdinal = lastReport.ordinal + 1;
  }

  // Set currency
  const currency = boutique.storeName === 'Belgrade' ? 'rsd' : 'eur';

  // Set reserved invoice number
  let invoiceNumber = '';
  if (boutique.storeName === 'Belgrade') {
    const currentYear = new Date().getFullYear().toString();
    let ordinalInvoice = `RID${currentYear.slice(-2)}`;
    let ordinalService = `RVD${currentYear.slice(-2)}`;
    const [[lastReportInv], [lastCheckoutSer]] = await Promise.all([
      Report.find({ store: boutique.store, invoiceNumber: { $regex: `.*${ordinalInvoice}.*` } })
        .sort({ invoiceNumber: -1 })
        .lean(),
      Checkout.find({ store: boutique.store, type: 'service', ordinal: { $regex: `.*${ordinalService}.*` } })
        .sort({ ordinal: -1 })
        .lean(),
    ]);
    let ordinalSuffix = '00001';
    if (lastReportInv || lastCheckoutSer) {
      let lastReportSuffix = lastReportInv ? lastReportInv.invoiceNumber.split('-') : ['0'];
      let lastCheckoutSuffix = lastCheckoutSer ? lastCheckoutSer.ordinal.split('-') : ['0'];
      lastReportSuffix = parseInt(lastReportSuffix[lastReportSuffix.length - 1]);
      lastCheckoutSuffix = parseInt(lastCheckoutSuffix[lastCheckoutSuffix.length - 1]);
      ordinalSuffix = Math.max(lastReportSuffix, lastCheckoutSuffix);
      ordinalSuffix += 1;
      ordinalSuffix = ordinalSuffix.toString();
      while (ordinalSuffix.length < 5) {
        ordinalSuffix = '0' + ordinalSuffix;
      }
    }
    invoiceNumber = `${ordinalInvoice}-${ordinalSuffix}`;
  }

  let purchaser = {};
  const toExecuteUpdate = [];
  const toExecuteSave = [];
  if (reservedForClient) {
    purchaser = {
      clientId: serialNumberObj.reservedFor,
      clientType: reservedForClient.clientType,
      firstName: reservedForClient.firstName,
      lastName: reservedForClient.lastName,
      fullName: reservedForClient.fullName,
      phones: reservedForClient.phones,
      emails: reservedForClient.emails,
      addresses: reservedForClient.addresses,
      quickNote: reservedForClient.quickNote,
      alias: reservedForClient.alias,
      passportNumber: reservedForClient.passportNumber,
    };

    // Update many wishlists -> set 'Recent purchase' status to client
    const projectionQuery = { 'clients.$': 1, entries: 1, archived: 1, product: 1, rmc: 1 };

    const [wishlists, productWishlist] = await Promise.all([
      Wishlist.find(
        {
          clients: { $elemMatch: { client: reservedForClient._id, status: { $in: onStatuses } } },
          product: { $ne: product._id },
        },
        projectionQuery
      )
        .populate('product')
        .lean(),
      Wishlist.findOne(
        {
          product: product._id,
          clients: { $elemMatch: { client: reservedForClient._id } },
        },
        projectionQuery
      ).lean(),
    ]);

    // Get sale product group
    const saleProductGroup = getProductGroup(product.brand, product.basicInfo.saleReference);

    for (const wishlist of wishlists) {
      const [clientOnWishlist] = wishlist.clients;

      let dateOfEnrollment = new Date();
      let adjustedDateOfEnrollmentComment = ` and the 'Adjusted date of enrollment' is set to today's date.`;
      if (['V', 'VI'].includes(saleProductGroup)) {
        dateOfEnrollment = clientOnWishlist.dateOfEnrollment;
        adjustedDateOfEnrollmentComment = '.';
      }

      // Set the date when the Recent purchase status changes to Active again
      const recentPurchaseProductGroup = getProductGroup(wishlist.product.brand, wishlist.product.basicInfo.saleReference);
      const recentPurchaseEnding = getRecentPurchaseValidation(recentPurchaseProductGroup, saleProductGroup, new Date());

      // Compile recent purchase ending date in the format DD.MM.YYYY
      const recentPurchaseEndingDay = new Date(recentPurchaseEnding.standbyEnding).getDate().toString();
      const recentPurchaseEndingMonth = (new Date(recentPurchaseEnding.standbyEnding).getMonth() + 1).toString();
      const recentPurchaseEndingYear = new Date(recentPurchaseEnding.standbyEnding).getFullYear().toString();
      const recentPurchaseEndingDate = `${recentPurchaseEndingDay}.${recentPurchaseEndingMonth}.${recentPurchaseEndingYear}`;

      // Create new activity
      let activityComment = `Bought the ${product.basicInfo.rmc}, on this wishlist the status is changed from '${clientOnWishlist.status}' to 'Recent purchase' until ${recentPurchaseEndingDate} (${recentPurchaseEnding.months} months)${adjustedDateOfEnrollmentComment}`;
      let setQuery = {
        'clients.$.status': 'Recent purchase',
        'clients.$.dateOfEnrollment': dateOfEnrollment,
        'clients.$.standbyEnding': recentPurchaseEnding.standbyEnding,
        'clients.$.previousStatus': clientOnWishlist.status,
      };

      if (clientOnWishlist.status === 'Recent purchase') {
        activityComment = `Bought the ${product.basicInfo.rmc}, on this wishlist 'Recent purchase' delayed until ${recentPurchaseEndingDate} (${recentPurchaseEnding.months} months)${adjustedDateOfEnrollmentComment}`;
        setQuery = {
          'clients.$.dateOfEnrollment': dateOfEnrollment,
          'clients.$.standbyEnding': recentPurchaseEnding.standbyEnding,
        };
      }

      const wishlistActivity = createActivity('Wishlist', createdBy, reservedForClient._id, activityComment, wishlist._id);

      // Push 'wishlistActivity' object in 'toExecuteSave' array
      toExecuteSave.push(wishlistActivity);

      // Update wishlist
      const updateWishlist = Wishlist.findOneAndUpdate(
        { _id: wishlist._id, clients: { $elemMatch: { client: reservedForClient._id, status: { $in: onStatuses } } } },
        {
          $set: setQuery,
          $push: { 'clients.$.logs': wishlistActivity._id },
        },
        { new: true }
      ).lean();

      // Push 'updateWishlist' object in 'toExecuteUpdate' array
      toExecuteUpdate.push(updateWishlist);

      // Update many shortlists -> set 'Recent purchase' status to client
      const updateManyShortlists = Shortlist.updateMany(
        { wishlist: wishlist._id, clients: { $elemMatch: { client: reservedForClient._id, status: { $in: onStatuses } } } },
        {
          $set: setQuery,
        },
        { new: true }
      ).lean();

      // Push 'updateManyShortlists' object in 'toExecuteUpdate' array
      toExecuteUpdate.push(updateManyShortlists);

      // Update client -> set new 'dateOfEnrollment' to client wishlist
      const updateClientWishlist = Client.findOneAndUpdate(
        { _id: reservedForClient._id, wishlists: { $elemMatch: { wishlist: wishlist._id } } },
        {
          $set: { 'wishlists.$.dateOfEnrollment': dateOfEnrollment },
        },
        { new: true }
      ).lean();

      // Push 'updateManyShortlists' object in 'toExecuteUpdate' array
      toExecuteUpdate.push(updateClientWishlist);
    }

    if (productWishlist) {
      const [newClientOnProductWishlist] = productWishlist.clients;

      // Create new activity
      activityComment = `Changed status from '${newClientOnProductWishlist.status}' to 'Watch sold' after purchase.`;
      const wishlistActivity = createActivity('Wishlist', createdBy, reservedForClient._id, activityComment, productWishlist._id);

      // Push 'wishlistActivity' object in 'toExecuteSave' array
      toExecuteSave.push(wishlistActivity);

      // Update wishlist -> set 'Watch sold' status to client
      const updateWishlist = Wishlist.findOneAndUpdate(
        { product: product._id, clients: { $elemMatch: { client: reservedForClient._id } } },
        {
          $set: {
            'clients.$.status': 'Watch sold',
            'clients.$.stopWaitingTime': new Date(),
            'clients.$.previousStatus': newClientOnProductWishlist.status,
          },
          $push: { 'clients.$.logs': wishlistActivity._id },
        },
        { new: true }
      ).lean();

      // Push 'updateWishlist' object in 'toExecuteUpdate' array
      toExecuteUpdate.push(updateWishlist);

      // Update shortlist -> set 'Watch sold' status to client
      const updateShortlist = Shortlist.updateMany(
        { wishlist: productWishlist._id, clients: { $elemMatch: { client: reservedForClient._id } } },
        {
          $set: {
            'clients.$.status': 'Watch sold',
            'clients.$.stopWaitingTime': new Date(),
            'clients.$.previousStatus': newClientOnProductWishlist.status,
          },
        },
        { new: true }
      ).lean();

      // Push 'updateShortlist' object in 'toExecuteUpdate' array
      toExecuteUpdate.push(updateShortlist);
    }
  }

  // Update product -> remove serial number object and create sales report
  const [updateProduct, salesReport] = await Promise.all([
    Product.findOneAndUpdate(
      { _id: productId, boutiques: { $elemMatch: { storeName: boutique.storeName } } },
      {
        $pull: { 'boutiques.$.serialNumbers': { number: serialNumberObj.number } },
        $inc: { 'boutiques.$.quantity': -1 },
      },
      { new: true }
    ).lean(),
    new Report({
      ordinal: reportOrdinal,
      status: 'requires data',
      language: 'sr',
      store: boutique.store,
      createdBy,
      responsiblePerson: createdBy,
      product,
      brand: product.brand,
      salesType: 'Sale',
      serialNumber: serialNumberObj,
      retailPrice: boutique.price,
      discount: 0,
      finalPrice: 0,
      vat: 0,
      netPrice: 0,
      invoicedLocalCurrency: 0,
      currency,
      payment: {
        card: 0,
        wireTransfer: 0,
        cash: 0,
      },
      salesDate: new Date(),
      manuallyCreated: true,
      invoiceNumber,
      purchaser,
    }).save(),
    toExecuteSave.map((model) => model.save()),
    ...toExecuteUpdate,
  ]);

  // Create activity
  await new Activity({
    type: 'Report',
    user: createdBy,
    comment: 'Manually declared as sold',
    product: productId,
    serialNumber: serialNumberObj.number,
    report: salesReport._id,
  }).save();

  return res.status(200).send({
    message: 'Successfully declared product as sold',
    results: salesReport,
  });
};

/**
 * @api {get} /product/excel/template Download Excel Template
 * @apiVersion 1.0.0
 * @apiName downloadTemplate
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String='Rolex', 'Tudor', 'Panerai', 'SwissKubik', 'Rubber B', 'Messika', 'Roberto Coin','Petrovic Diamonds', 'RolexRMC', 'TudorRMC', 'PaneraiRMC', 'SwissKubikRMC', 'RubberBRMC'} brand Brand Type
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
 {
   "message": "Successfully downloaded Excel template",
 }
 * @apiUse MissingParamsError
 * @apiUse InvalidValue
 */
 module.exports.downloadTemplate = async (req, res) => {
  let { brand } = req.query;

  const brands = ['Rolex', 'Tudor', 'Panerai', 'SwissKubik', 'Rubber B', 'Messika', 'Roberto Coin','Petrovic Diamonds', 'RolexRMC', 'TudorRMC', 'PaneraiRMC', 'SwissKubikRMC', 'RubberBRMC'];

  if (!brand) throw new Error(error.MISSING_PARAMETERS);
  if (!brands.includes(brand)) throw new Error(error.INVALID_VALUE);

  // Create excel workbook
  const workbook = new exceljs.Workbook();

  // Create sheet
  const worksheet = workbook.addWorksheet(`Import ${brand}`);
  worksheet.addRow();

  // Set font
  const font = { name: 'Arial', size: 12 };

  // Switch brands
  switch (brand) {
    // ROLEX and TUDOR
    case 'Rolex':
    case 'Tudor':
      // Set columns
      worksheet.columns = [
        {
          header: 'PGP',
          key: 'pgpReference',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Location',
          key: 'location',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Status',
          key: 'status',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Reserved For - Client ID',
          key: 'reservedFor',
          width: 26,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Reservation Time',
          key: 'reservationTime',
          width: 18,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Comment',
          key: 'comment',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Previous Serial',
          key: 'serialNumber',
          width: 17,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'soldToParty',
          key: 'soldToParty',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'shipToParty',
          key: 'shipToParty',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'billToParty',
          key: 'billToParty',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'packingNumber',
          key: 'packingNumber',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'invoiceDate',
          key: 'invoiceDate',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'invoiceNumber',
          key: 'invoiceNumber',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Serial',
          key: 'serialNumber2',
          width: 17,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'boxCode',
          key: 'boxCode',
          width: 12,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'exGenevaPrice',
          key: 'exGenevaPrice',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'sectorId',
          key: 'sectorId',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'sector',
          key: 'sector',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'boitId',
          key: 'boitId',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'boitRef',
          key: 'boitRef',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'caspId',
          key: 'caspId',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'cadranDesc',
          key: 'cadranDesc',
          width: 18,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'ldisId',
          key: 'ldisId',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'ldisq',
          key: 'ldisq',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'brspId',
          key: 'brspId',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'brspRef',
          key: 'brspRef',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'prixSuisse',
          key: 'prixSuisse',
          width: 12,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'RMC',
          key: 'rmc',
          width: 18,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'umIntern',
          key: 'umIntern',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'umExtern',
          key: 'umExtern',
          width: 11,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'rfid',
          key: 'rfid',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'packingType',
          key: 'packingType',
          width: 14,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Shipment Date (YYYY-MM-DD)',
          key: 'shipmentDate',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        }
      ];

      // Add rows
      for (let i = 0; i < 30; i++) {
        let row = worksheet.addRow({ pgpReference: '' });
      }

      // Design columns by column names
      for (let i = 1; i < 34; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'E5DAC4' }         // LIGHT GOLD
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },          // WHITE
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      for (let i = 8; i < 33; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'D9E7DF' }         // LIGHT GREEN
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      for (let i = 1; i < 4; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }       // DARK GOLD
          };
        });
      }

      worksheet.getColumn(7).eachCell((cell, rowNumber) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'CFBC93' }         // DARK GOLD
        };
      });

      worksheet.getColumn(14).eachCell((cell, rowNumber) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'B2CFC3' }        // MIDDLE GREEN
        };
      });

      worksheet.getColumn(16).eachCell((cell, rowNumber) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'B2CFC3' }        // MIDDLE GREEN
        };
      });

      worksheet.getColumn(28).eachCell((cell, rowNumber) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'B2CFC3' }         // MIDDLE GREEN
        };
      });

      worksheet.getColumn(33).eachCell((cell, rowNumber) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'CFBC93' }         // DARK GOLD
        };
      });

      // Design header row
      for(let i = 1; i < 34; i++) {
        worksheet.getRow(1).getCell(i).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '006039' },        // DARK GREEN
        };
        worksheet.getRow(1).getCell(i).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };        // WHITE
        worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getRow(1).getCell(i).border = {
          top: { style: 'thin', color: { argb: 'FFFFFF' } },
          left: { style: 'thin', color: { argb: 'FFFFFF' } },
          bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
          right: { style: 'thin', color: { argb: 'FFFFFF' } },
        };
      }
      break;

      // PANERAI
      case 'Panerai':
        // Set columns
        worksheet.columns = [
          {
            header: 'Sale Reference',
            key: 'saleReference',
            width: 18,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Serial Number',
            key: 'serialNumber',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'PGP',
            key: 'pgpReference',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Location',
            key: 'location',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Comment',
            key: 'comment',
            width: 30,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Stock Date',
            key: 'stockDate',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Date',
            key: 'invoiceDate',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Purchase Price (RSD/HUF)',
            key: 'purchasePriceLocal',
            width: 26,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          }
        ];

        // Add rows
        for (let i = 0; i < 20; i++) {
          let row = worksheet.addRow({ pgpReference: '' });
        }

        // Design all columns
        for (let i = 1; i < 9; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'E5DAC4' }         // LIGHT GOLD
            };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FFFFFF' } },        // WHITE
              left: { style: 'thin', color: { argb: 'FFFFFF' } },
              bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
              right: { style: 'thin', color: { argb: 'FFFFFF' } },
            };
          });
        }

        //  Design first four columns
        for (let i = 1; i < 5; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        // Design header row
        for(let i = 1; i < 9; i++) {
          worksheet.getRow(1).getCell(i).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '006039' },          // DARK GREEN
          };
          worksheet.getRow(1).getCell(i).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
          worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
          worksheet.getRow(1).getCell(i).border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        }
        break;

      // SWISSKUBIK
      case 'SwissKubik':
        // Set columns
        worksheet.columns = [
          {
            header: 'Sale Reference',
            key: 'saleReference',
            width: 18,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Serial Number',
            key: 'serialNumber',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'PGP',
            key: 'pgpReference',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Location',
            key: 'location',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Comment',
            key: 'comment',
            width: 30,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Stock Date',
            key: 'stockDate',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Number',
            key: 'invoiceNumber',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Date',
            key: 'invoiceDate',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Ex Geneva Price (*100)',
            key: 'exGenevaPrice',
            width: 24,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          },
          {
            header: 'Purchase Price (RSD/HUF)',
            key: 'purchasePriceLocal',
            width: 26,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          }
        ];

        // Add rows
        for (let i = 0; i < 20; i++) {
          let row = worksheet.addRow({ pgpReference: '' });
        }

        // Design all clomuns
        for (let i = 1; i < 11; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'E5DAC4' }         // LIGHT GOLD
            };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FFFFFF' } },        // WHITE
              left: { style: 'thin', color: { argb: 'FFFFFF' } },
              bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
              right: { style: 'thin', color: { argb: 'FFFFFF' } },
            };
          });
        }

        // Design first four columns
        for (let i = 1; i < 5; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        // Design 'invoiceDate' and 'exGenevaPrice' columns
        for (let i = 8; i < 10; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        // Design header row
        for(let i = 1; i < 11; i++) {
          worksheet.getRow(1).getCell(i).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '006039' },          // DARK GREEN
          };
          worksheet.getRow(1).getCell(i).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
          worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
          worksheet.getRow(1).getCell(i).border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        }
        break;

      // RUBBER B
      case 'Rubber B':
        // Set columns
        worksheet.columns = [
          {
            header: 'RMC',
            key: 'saleReference',
            width: 18,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Serial Number',
            key: 'serialNumber',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'PGP',
            key: 'pgpReference',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Location',
            key: 'location',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Size',
            key: 'size',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Comment',
            key: 'comment',
            width: 30,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Stock Date',
            key: 'stockDate',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Date',
            key: 'invoiceDate',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Purchase Price (RSD/HUF)',
            key: 'purchasePriceLocal',
            width: 26,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          }
        ];

        // Add rows
        for (let i = 0; i < 20; i++) {
          let row = worksheet.addRow({ pgpReference: '' });
        }

        // Design all columns
        for (let i = 1; i < 10; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'E5DAC4' }         // LIGHT GOLD
            };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FFFFFF' } },        // WHITE
              left: { style: 'thin', color: { argb: 'FFFFFF' } },
              bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
              right: { style: 'thin', color: { argb: 'FFFFFF' } },
            };
          });
        }

        // Design first five columns
        for (let i = 1; i < 6; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        // Design header row
        for(let i = 1; i < 10; i++) {
          worksheet.getRow(1).getCell(i).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '006039' },          // DARK GREEN
          };
          worksheet.getRow(1).getCell(i).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
          worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
          worksheet.getRow(1).getCell(i).border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        }
        break;

      // MESSIKA
      case 'Messika':
        // Set columns
        worksheet.columns = [
          {
            header: 'Collection',
            key: 'rmc',
            width: 14,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Sale Reference',
            key: 'saleReference',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Material Description',
            key: 'materialDescription',
            width: 26,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Jewelry Type',
            key: 'jewelryType',
            width: 14,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Size',
            key: 'size',
            width: 8,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Gold Weight',
            key: 'weight',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Stones Weight',
            key: 'allStonesWeight',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Stones Qty',
            key: 'stonesQty',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Brilliants',
            key: 'brilliants',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Purchase Price',
            key: 'purchasePrice',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Purchase Price (RSD/HUF)',
            key: 'purchasePriceLocal',
            width: 26,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          },
          {
            header: 'Retail Price',
            key: 'price',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Serial Number',
            key: 'serialNumber',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'PGP',
            key: 'pgpReference',
            width: 10,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Location',
            key: 'location',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Comment',
            key: 'comment',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Stock Date',
            key: 'stockDate',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Number',
            key: 'invoiceNumber',
            width: 17,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Date',
            key: 'invoiceDate',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          }
        ];

        // Add rows
        for (let i = 0; i < 20; i++) {
          let row = worksheet.addRow({ pgpReference: '' });
        }

        // Design all columns
        for (let i = 1; i < 20; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'E5DAC4' }         // LIGHT GOLD
            };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FFFFFF' } },        // WHITE
              left: { style: 'thin', color: { argb: 'FFFFFF' } },
              bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
              right: { style: 'thin', color: { argb: 'FFFFFF' } },
            };
          });
        }

        // Design first four columns
        for (let i = 1; i < 5; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        // Design 'purchasePrice' column
        worksheet.getColumn(10).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }         // DARK GOLD
          };
        });

        // Design 'retailPrice' column
        worksheet.getColumn(12).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'B2CFC3' }       // MIDDLE GREEN
          };
        });

        // Design serial number columns
        for (let i = 13; i < 16; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        // Design header row
        for(let i = 1; i < 20; i++) {
          worksheet.getRow(1).getCell(i).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '006039' },          // DARK GREEN
          };
          worksheet.getRow(1).getCell(i).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
          worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
          worksheet.getRow(1).getCell(i).border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        }
        break;

      // ROBERTO COIN
      case 'Roberto Coin':
        // Set columns
        worksheet.columns = [
          {
            header: 'Collection',
            key: 'rmc',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Sale Reference',
            key: 'saleReference',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Jewelry Type',
            key: 'jewelryType',
            width: 14,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Size',
            key: 'size',
            width: 8,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Gold Color',
            key: 'materials',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Gold Weight',
            key: 'weight',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia Gia',
            key: 'diaGia',
            width: 9,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia Qty',
            key: 'diaQty',
            width: 9,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia Carat',
            key: 'diaCarat',
            width: 11,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Ruby Qty',
            key: 'rubyQty',
            width: 10,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Ruby Carat',
            key: 'rubyCarat',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Stones',
            key: 'brilliants',
            width: 20,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Purchase Price',
            key: 'purchasePrice',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          },
          {
            header: 'Purchase Price (RSD/HUF)',
            key: 'purchasePriceLocal',
            width: 26,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          },
          {
            header: 'Retail Price',
            key: 'price',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          },
          {
            header: 'Serial Number',
            key: 'serialNumber',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'PGP',
            key: 'pgpReference',
            width: 10,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Location',
            key: 'location',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Comment',
            key: 'comment',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Stock Date',
            key: 'stockDate',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Number',
            key: 'invoiceNumber',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Date',
            key: 'invoiceDate',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          }
        ];

        // Add rows
        for (let i = 0; i < 20; i++) {
          let row = worksheet.addRow({ pgpReference: '' });
        }

        // Design all columns
        for (let i = 1; i < 23; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'E5DAC4' }         // LIGHT GOLD
            };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FFFFFF' } },        // WHITE
              left: { style: 'thin', color: { argb: 'FFFFFF' } },
              bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
              right: { style: 'thin', color: { argb: 'FFFFFF' } },
            };
          });
        }

        //  Design first three columns
        for (let i = 1; i < 4; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        //  Design 'purchasePrice' column
        worksheet.getColumn(13).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }         // DARK GOLD
          };
        });

        // Design 'retailPrice' column
        worksheet.getColumn(15).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'B2CFC3' }       // MIDDLE GREEN
          };
        });

        // Design serial number columns
        for (let i = 16; i < 19; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        // Design header row
        for(let i = 1; i < 23; i++) {
          worksheet.getRow(1).getCell(i).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '006039' },          // DARK GREEN
          };
          worksheet.getRow(1).getCell(i).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
          worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
          worksheet.getRow(1).getCell(i).border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        }
        break;

      // PETROVIC DIAMONDS
      case 'Petrovic Diamonds':
        // Set columns
        worksheet.columns = [
          {
            header: 'Sale Reference',
            key: 'saleReference',
            width: 18,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Jewelry Type',
            key: 'jewelryType',
            width: 14,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Size',
            key: 'size',
            width: 8,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Gold Weight',
            key: 'weight',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 1 Carat',
            key: 'dia1Carat',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 1 Color',
            key: 'dia1Color',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 1 Clarity',
            key: 'dia1Clarity',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 1 Shape',
            key: 'dia1Shape',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 1 Cut',
            key: 'dia1Cut',
            width: 10,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 1 Polish',
            key: 'dia1Polish',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 1 Symmetry',
            key: 'dia1Symmetry',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 1 Certificate',
            key: 'dia1Certificate',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 2 Carat',
            key: 'dia2Carat',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 2 Color',
            key: 'dia2Color',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 2 Clarity',
            key: 'dia2Clarity',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 2 Shape',
            key: 'dia2Shape',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 2 Cut',
            key: 'dia2Cut',
            width: 10,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 2 Polish',
            key: 'dia2Polish',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 2 Symmetry',
            key: 'dia2Symmetry',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 2 Certificate',
            key: 'dia2Certificate',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 3 Carat',
            key: 'dia3Carat',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 3 Color',
            key: 'dia3Color',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 3 Clarity',
            key: 'dia3Clarity',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Dia 3 Shape',
            key: 'dia3Shape',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 3 Cut',
            key: 'dia3Cut',
            width: 10,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 3 Polish',
            key: 'dia3Polish',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 3 Symmetry',
            key: 'dia3Symmetry',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Dia 3 Certificate',
            key: 'dia3Certificate',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Purchase Price (RSD/HUF)',
            key: 'purchasePriceLocal',
            width: 26,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          },
          {
            header: 'Retail Price',
            key: 'price',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
          },
          {
            header: 'Serial Number',
            key: 'serialNumber',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'PGP',
            key: 'pgpReference',
            width: 10,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Location',
            key: 'location',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Comment',
            key: 'comment',
            width: 15,
            style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
          },
          {
            header: 'Stock Date',
            key: 'stockDate',
            width: 12,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Number',
            key: 'invoiceNumber',
            width: 16,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          },
          {
            header: 'Invoice Date',
            key: 'invoiceDate',
            width: 13,
            style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
          }
        ];

        // Add rows
        for (let i = 0; i < 20; i++) {
          let row = worksheet.addRow({ pgpReference: '' });
        }

        // Design all columns
        for (let i = 1; i < 38; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'E5DAC4' }         // LIGHT GOLD
            };
            cell.border = {
              top: { style: 'thin', color: { argb: 'FFFFFF' } },        // WHITE
              left: { style: 'thin', color: { argb: 'FFFFFF' } },
              bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
              right: { style: 'thin', color: { argb: 'FFFFFF' } },
            };
          });
        }

        // Design first two columns
        for (let i = 1; i < 3; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        // Design 'retailPrice' column
        for (let i = 30; i < 31; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'B2CFC3' }       // MIDDLE GREEN
            };
          });
        }

        // Design serial number columns
        for (let i = 31; i < 34; i++) {
          worksheet.getColumn(i).eachCell((cell, rowNumber) => {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'CFBC93' }         // DARK GOLD
            };
          });
        }

        // Design header row
        for(let i = 1; i < 38; i++) {
          worksheet.getRow(1).getCell(i).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '006039' },          // DARK GREEN
          };
          worksheet.getRow(1).getCell(i).font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFF' } };
          worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
          worksheet.getRow(1).getCell(i).border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        }
        break;

    // ROLEX RMC
    case 'RolexRMC':
      // Set columns
      worksheet.columns = [
        {
          header: 'RMC',
          key: 'rmc',
          width: 18,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Collection',
          key: 'collection',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Product Line',
          key: 'productLine',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Sale Reference',
          key: 'saleReference',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Material Description',
          key: 'materialDescription',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Dial',
          key: 'dial',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Bracelet',
          key: 'bracelet',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Box',
          key: 'box',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Diameter',
          key: 'diameter',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Case Material',
          key: 'caseMaterial',
          width: 14,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Ex_Geneve_CHF',
          key: 'exGeneveCHF',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'Retail_EUR_RS',
          key: 'retailRsEUR',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'Retail_EUR_HU',
          key: 'retailHuEUR',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'Retail_EUR_MNE',
          key: 'retailMneEUR',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        }
      ];

      // Add ROWS
      for (let i = 0; i < 20; i++) {
        let row = worksheet.addRow({ rmc: '' });
      }

      // Design all COLUMNS besides prices
      for (let i = 1; i < 12; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'E5DAC4' }     // LIGHT GOLD
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },      // WHITE
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design price COLUMNS
      for (let i = 12; i < 15; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'B2CFC3' }       // MIDDLE GREEN
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design required COLUMNS
      for (let i = 1; i < 3; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }       // DARK GOLD
          };
        });
      }

      for (let i = 4; i < 6; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }       // DARK GOLD
          };
        });
      }

      // Design HEADER ROW
      for(let i = 1; i < 15; i++) {
        worksheet.getRow(1).getCell(i).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '006039' },      // DARK GREEN
        };
        worksheet.getRow(1).getCell(i).font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getRow(1).getCell(i).border = {
          top: { style: 'thin', color: { argb: 'FFFFFF' } },
          left: { style: 'thin', color: { argb: 'FFFFFF' } },
          bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
          right: { style: 'thin', color: { argb: 'FFFFFF' } },
        };
      }
      break;

    // TUDOR RMC
    case 'TudorRMC':
      // Set columns
      worksheet.columns = [
        {
          header: 'RMC',
          key: 'rmc',
          width: 18,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Collection',
          key: 'collection',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Product Line',
          key: 'productLine',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Sale Reference',
          key: 'saleReference',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Material Description',
          key: 'materialDescription',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Dial',
          key: 'dial',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Bracelet',
          key: 'bracelet',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Box',
          key: 'box',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Diameter',
          key: 'diameter',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Case Material',
          key: 'caseMaterial',
          width: 14,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Waterproofness',
          key: 'waterproofness',
          width: 14,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Ex_Geneve_CHF',
          key: 'exGeneveCHF',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'Retail_EUR_RS',
          key: 'retailRsEUR',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        }
      ];

      // Add ROWS
      for (let i = 0; i < 20; i++) {
        let row = worksheet.addRow({ rmc: '' });
      }

      // Design all COLUMNS besides prices
      for (let i = 1; i < 13; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'E5DAC4' }     // LIGHT GOLD
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },      // WHITE
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design price COLUMNS
      for (let i = 13; i < 14; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'B2CFC3' }       // MIDDLE GREEN
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design required COLUMNS
      for (let i = 1; i < 3; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }       // DARK GOLD
          };
        });
      }

      for (let i = 4; i < 6; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }       // DARK GOLD
          };
        });
      }

      // Design HEADER ROW
      for(let i = 1; i < 14; i++) {
        worksheet.getRow(1).getCell(i).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '006039' },        // DARK GREEN
        };
        worksheet.getRow(1).getCell(i).font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getRow(1).getCell(i).border = {
          top: { style: 'thin', color: { argb: 'FFFFFF' } },
          left: { style: 'thin', color: { argb: 'FFFFFF' } },
          bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
          right: { style: 'thin', color: { argb: 'FFFFFF' } },
        };
      }
      break;

    // PANERAI RMC
    case 'PaneraiRMC':
      // Set columns
      worksheet.columns = [
        {
          header: 'RMC',
          key: 'rmc',
          width: 18,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Collection',
          key: 'collection',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Sale Reference',
          key: 'saleReference',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Material Description',
          key: 'materialDescription',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Dial',
          key: 'dial',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Bracelet',
          key: 'bracelet',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Movement',
          key: 'movement',
          width: 13,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Diameter',
          key: 'diameter',
          width: 10,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Case Material',
          key: 'caseMaterial',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Purchase Price',
          key: 'purchasePrice',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'Retail_EUR_RS',
          key: 'retailRsEUR',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        }
      ];

      // Add ROWS
      for (let i = 0; i < 20; i++) {
        let row = worksheet.addRow({ rmc: '' });
      }

      // Design all COLUMNS besides prices
      for (let i = 1; i < 11; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'E5DAC4' }     // LIGHT GOLD
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },      // WHITE
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design price COLUMNS
      for (let i = 11; i < 12; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'B2CFC3' }       // MIDDLE GREEN
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design required COLUMNS
      for (let i = 1; i < 4; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }       // DARK GOLD
          };
        });
      }

      worksheet.getColumn(10).eachCell((cell, rowNumber) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'CFBC93' }       // DARK GOLD
        };
      });

      // Design HEADER ROW
      for(let i = 1; i < 12; i++) {
        worksheet.getRow(1).getCell(i).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '006039' },        // DARK GREEN
        };
        worksheet.getRow(1).getCell(i).font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getRow(1).getCell(i).border = {
          top: { style: 'thin', color: { argb: 'FFFFFF' } },
          left: { style: 'thin', color: { argb: 'FFFFFF' } },
          bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
          right: { style: 'thin', color: { argb: 'FFFFFF' } },
        };
      }
      break;

    // SWISSKUBIK RMC
    case 'SwissKubikRMC':
      // Set columns
      worksheet.columns = [
        {
          header: 'RMC',
          key: 'rmc',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Collection',
          key: 'collection',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Sale Reference',
          key: 'saleReference',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Material Description',
          key: 'materialDescription',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Description',
          key: 'description',
          width: 30,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Size',
          key: 'size',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Materials',
          key: 'materials',
          width: 25,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Color',
          key: 'color',
          width: 12,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Retail_EUR_RS',
          key: 'retailRsEUR',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'Retail_EUR_HU',
          key: 'retailHuEUR',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        }
      ];

      // Add ROWS
      for (let i = 0; i < 20; i++) {
        let row = worksheet.addRow({ rmc: '' });
      }

      // Design all COLUMNS besides prices
      for (let i = 1; i < 10; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'E5DAC4' }     // LIGHT GOLD
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },      // WHITE
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design price COLUMNS
      for (let i = 9; i < 11; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'B2CFC3' }       // MIDDLE GREEN
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design required COLUMNS
      for (let i = 1; i < 4; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }       // DARK GOLD
          };
        });
      }

      // Design HEADER ROW
      for(let i = 1; i < 11; i++) {
        worksheet.getRow(1).getCell(i).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '006039' },        // DARK GREEN
        };
        worksheet.getRow(1).getCell(i).font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getRow(1).getCell(i).border = {
          top: { style: 'thin', color: { argb: 'FFFFFF' } },
          left: { style: 'thin', color: { argb: 'FFFFFF' } },
          bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
          right: { style: 'thin', color: { argb: 'FFFFFF' } },
        };
      }
      break;

    // RUBBER B RMC
    case 'RubberBRMC':
      // Set columns
      worksheet.columns = [
        {
          header: 'RMC',
          key: 'rmc',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Sale Reference',
          key: 'saleReference',
          width: 15,
          style: { font, alignment: { vertical: 'middle', horizontal: 'center' } },
        },
        {
          header: 'Material Description',
          key: 'materialDescription',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Color',
          key: 'color',
          width: 20,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'For Model',
          key: 'forModel',
          width: 25,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'For Clasp',
          key: 'forClasp',
          width: 25,
          style: { font, alignment: { vertical: 'middle', horizontal: 'left' } },
        },
        {
          header: 'Purchase Price',
          key: 'purchasePrice',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'Retail_EUR_RS',
          key: 'retailRsEUR',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        },
        {
          header: 'Retail_EUR_HU',
          key: 'retailHuEUR',
          width: 16,
          style: { font, alignment: { vertical: 'middle', horizontal: 'right' } },
        }
      ];

      // Add ROWS
      for (let i = 0; i < 20; i++) {
        let row = worksheet.addRow({ rmc: '' });
      }

      // Design all COLUMNS besides prices
      for (let i = 1; i < 8; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'E5DAC4' }     // LIGHT GOLD
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },      // WHITE
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design price COLUMNS
      for (let i = 8; i < 10; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'B2CFC3' }       // MIDDLE GREEN
          };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FFFFFF' } },
            left: { style: 'thin', color: { argb: 'FFFFFF' } },
            bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
            right: { style: 'thin', color: { argb: 'FFFFFF' } },
          };
        });
      }

      // Design required COLUMNS
      for (let i = 1; i < 3; i++) {
        worksheet.getColumn(i).eachCell((cell, rowNumber) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'CFBC93' }       // DARK GOLD
          };
        });
      }

      worksheet.getColumn(7).eachCell((cell, rowNumber) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'CFBC93' }       // DARK GOLD
        };
      });

      // Design HEADER ROW
      for(let i = 1; i < 10; i++) {
        worksheet.getRow(1).getCell(i).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '006039' },        // DARK GREEN
        };
        worksheet.getRow(1).getCell(i).font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(1).getCell(i).alignment = { vertical: 'middle', horizontal: 'center' };
        worksheet.getRow(1).getCell(i).border = {
          top: { style: 'thin', color: { argb: 'FFFFFF' } },
          left: { style: 'thin', color: { argb: 'FFFFFF' } },
          bottom: { style: 'thin', color: { argb: 'FFFFFF' } },
          right: { style: 'thin', color: { argb: 'FFFFFF' } },
        };
      }
      break;
  }

  worksheet.getRow(1).height = 26;

  const createDir = util.promisify(tmp.dir);
  const tmpDir = await createDir();
  const filePath = `${tmpDir}/${uuidv4()}.xlsx`;
  // const filePath = `/Users/aleksandar/desktop/${Date.now()}.xlsx`;

  return workbook.xlsx.writeFile(filePath).then(() => {
    const stream = fs.createReadStream(filePath);

    stream.on('error', () => {
      throw new Error(error.BAD_REQUEST);
    });
    stream.pipe(res);
  });
};

/**
 * @api {delete} /product/soonInStock Delete soon in stock products
 * @apiVersion 1.0.0
 * @apiName deleteSoonInStocks
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {String[]} soonInStockId SoonInStock ID
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 201 OK
 {
    "message": "Successfully deleted soon in stock watches"
 }
 * @apiUse MissingParamsError
 * @apiUse NotFound
 * @apiUse CredentialsError
 */
 module.exports.deleteSoonInStocks = async (req, res) => {
  const { soonInStockId } = req.query;

  if (!soonInStockId) throw new Error(error.MISSING_PARAMETERS);

  const results = await SoonInStock.deleteMany({ _id: { $in: soonInStockId } });

  if (results.n === 0) throw new Error(error.NOT_FOUND);

  return res.status(200).send({
    message: 'Successfully deleted soon in stock watches',
  });
};

/**
 * @api {get} /product/labels/pgp-reference Get products by pgp reference range
 * @apiVersion 1.0.0
 * @apiName Get products by pgp reference range
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (query) {Number} [pgpReferenceFrom] Pgp reference from
 * @apiParam (query) {Number} [pgpReferenceTo] Pgp reference to
 * @apiParam (query) {String} [pgpReference] Pgp reference, NOTE: ?pgpReference=123&pgpReference=456
 * @apiParam (query) {String="pgpReference","-pgpReference"} [sortBy] Sort products by pgp reference
 * @apiParam (query) {String} [brand] Brand name
 * @apiParam (query) {String} [store] Store ID
 * @apiParam (query) {String} [serialNumberBelgrade] Serial number for all labels for products from Belgrade
 * @apiParam (query) {String} [serialNumberBudapest] Serial number for all labels for products from Budapest
 * @apiParam (query) {String} [serialNumberMontenegro] Serial number for all labels for products from Montenegro
 * @apiParam (query) {String} [soonInStock='true'] Include soon in stock products
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
{
  "message": "Products with pgp references successfully fetched",
  "results": [
    {
      "_id": "603caa8ac4ef701214fae3ca",
      "basicInfo": {
        "rmc": "ADR777EA2781",
        "collection": "Palazzo Ducale",
        "saleReference": "ADR777EA2781",
        "jewelryType": "earrings",
        "purchasePrice": 559,
        "materials": [
          "Rose gold",
        ],
        "weight": "3.82",
        "stonesQty": 8,
        "allStonesWeight": 0.156,
        "stones": [
          {
            "_id": "6087bb6d2f4b24766f8024b1",
            "type": "Ruby",
            "quantity": 1,
            "stoneTypeWeight": "0.036",
          },
        ],
        "brilliants": "",
        "diaGia": "GVS",
        "photos": [
          "ADR777EA2781.jpg",
          "ADR777EA2781 copy-LR-1200-1616971983800.jpg",
        ],
        "productLine": "",
        "diamonds": [],
      },
      "quotaRegime": false,
      "status": "previous",
      "brand": "Roberto Coin",
      "boutiques": {
        "quantity": 1,
        "_id": "6087bb6d2f4b24766f8024b2",
        "store": "5f3a4225ffe375404f72fb06",
        "storeName": "Belgrade",
        "price": 1350,
        "priceLocal": 162000,
        "VATpercent": 20,
        "priceHistory": [
          {
            "_id": "6087bb6d2f4b24766f8024b3",
            "date": "2020-01-01T00:00:00.000Z",
            "price": 1350,
            "VAT": 20,
            "priceLocal": 162000,
          },
        ],
        "serialNumbers": {
          "warrantyConfirmed": false,
          "modified": false,
          "_id": "6087bb6d2f4b24766f8024b4",
          "number": "10008",
          "pgpReference": "10008",
          "status": "Stock",
          "location": "Drawer Sales Floor - MBD",
          "stockDate": "2020-12-15T23:00:00.000Z",
          "comment": "",
          "pgpReferenceNumber": 10008,
          "labelText": "R10673 - ROLEX\n114060 - Oyster - M\nBLACK INDEX W\n1/A120 - 7450 - 894'000 RSD",
          "declaration": "R10673 - DEKLARACIJA: Ručni časovnik ROLEX Materijal: Čelik i plemeniti metali Zemlja porekla: CH Uvoznik: Petite Genève Petrović doo, Uskočka 7, Beograd",
          "qrCodeLink": "http://deklaracija.pgp.rs/deklaracija?brand=rolex::SUBMARINER:114060"
        },
      },
      "__v": 0,
      "createdAt": "2021-03-01T08:49:16.546Z",
      "updatedAt": "2021-07-28T13:30:00.189Z",
      "active": true,
    },
  ],
}
 * @apiUse InvalidValue
 * @apiUse CredentialsError
 */
module.exports.getProductsByPgpRange = async (req, res) => {

  let { pgpReferenceFrom, pgpReferenceTo, pgpReference, sortBy, store, brand, serialNumberBelgrade, serialNumberBudapest, serialNumberMontenegro, soonInStock = 'true' } = req.query;

  const aggregatePipelineProducts = [
    { $unwind: '$boutiques' },
    { $unwind: '$boutiques.serialNumbers' },
    {
      $addFields:
      {
        'boutiques.serialNumbers.pgpReferenceNumber': {
          $convert: {
            input: '$boutiques.serialNumbers.pgpReference', to: 'int', onError: null, onNull: null,
          },
        },
      },
    },
  ];

  const aggregatePipelineSoonInStock = [
    {
      $addFields:
      {
        pgpReferenceNumber: {
          $convert: {
            input: '$pgpReference', to: 'int', onError: null, onNull: null,
          },
        },
      },
    },
    {
      $lookup:
        {
          from: 'products',
          localField: 'product',
          foreignField: '_id',
          as: 'product'
        }
    },
    { $unwind: '$product' },
  ];

  // Handle store
  if (store) {
    if (!isValidId(store)) throw new Error(error.INVALID_VALUE);
    aggregatePipelineProducts.push({ $match: { 'boutiques.store': ObjectId(store) } });
    aggregatePipelineSoonInStock.push({ $match: { store: ObjectId(store) } });
  }

  if (brand) {
    if (!Array.isArray(brand)) brand = [brand];
    aggregatePipelineProducts.push({ $match: { brand: { $in: brand } } });
    aggregatePipelineSoonInStock.push({ $match: { brand: { $in: brand } } });
  }

  // Sort handler
  if (sortBy) {
    const allowedValues = ['pgpReference', '-pgpReference'];
    if (!allowedValues.includes(sortBy)) throw new Error(error.INVALID_VALUE);
    aggregatePipelineProducts.push({ $sort: { 'boutiques.serialNumbers.pgpReferenceNumber': sortBy.charAt(0) === '-' ? -1 : 1 } });
    aggregatePipelineSoonInStock.push({ $sort: { pgpReferenceNumber: sortBy.charAt(0) === '-' ? -1 : 1 } });
  }

  // Pgp range handler
  let rangeExpressionProducts;
  let rangeExpressionSoonInStock;
  pgpReferenceFrom = +pgpReferenceFrom;
  pgpReferenceTo = +pgpReferenceTo;
  if (pgpReferenceFrom && pgpReferenceTo) {
    if (!Number.isInteger(pgpReferenceFrom)) throw new Error(error.INVALID_VALUE);
    if (!Number.isInteger(pgpReferenceTo)) throw new Error(error.INVALID_VALUE);
    rangeExpressionProducts = { 'boutiques.serialNumbers.pgpReferenceNumber': { $lte: pgpReferenceTo, $gte: pgpReferenceFrom } };
    rangeExpressionSoonInStock = { pgpReferenceNumber: { $lte: pgpReferenceTo, $gte: pgpReferenceFrom } };
  } else if (pgpReferenceFrom) {
    if (!Number.isInteger(pgpReferenceFrom)) throw new Error(error.INVALID_VALUE);
    rangeExpressionProducts = { 'boutiques.serialNumbers.pgpReferenceNumber': { $gte: pgpReferenceFrom } };
    rangeExpressionSoonInStock = { pgpReferenceNumber: { $gte: pgpReferenceFrom } };
  } else if (pgpReferenceTo) {
    if (!Number.isInteger(pgpReferenceTo)) throw new Error(error.INVALID_VALUE);
    rangeExpressionProducts = { 'boutiques.serialNumbers.pgpReferenceNumber': { $lte: pgpReferenceTo } };
    rangeExpressionSoonInStock = { pgpReferenceNumber: { $lte: pgpReferenceTo } };
  }

  // Pgp references handler
  let matchExpressionProducts;
  let matchExpressionSoonInStock;
  if (pgpReference) {
    if (!Array.isArray(pgpReference)) pgpReference = [pgpReference];
    if (rangeExpressionProducts) {
      matchExpressionProducts = { $match: { $or: [rangeExpressionProducts, { 'boutiques.serialNumbers.pgpReference': { $in: pgpReference } }] } };
    } else {
      matchExpressionProducts = { $match: { 'boutiques.serialNumbers.pgpReference': { $in: pgpReference } } };
    }
    if (rangeExpressionSoonInStock) {
      matchExpressionSoonInStock = { $match: { $or: [rangeExpressionSoonInStock, { pgpReference: { $in: pgpReference } }] } };
    } else {
      matchExpressionSoonInStock = { $match: { pgpReference: { $in: pgpReference } } };
    }
  } else {
    if (rangeExpressionProducts) {
      matchExpressionProducts = { $match: rangeExpressionProducts };
    }
    if (rangeExpressionSoonInStock) {
      matchExpressionSoonInStock = { $match: rangeExpressionSoonInStock };
    }
  }

  if (matchExpressionProducts) aggregatePipelineProducts.push(matchExpressionProducts);
  if (matchExpressionSoonInStock) aggregatePipelineSoonInStock.push(matchExpressionSoonInStock);

  const cursorProducts = Product.aggregate(aggregatePipelineProducts);

  // check serial number
  const labelSerialsMap = new Map();
  const [stores, labelSerialsDocs] = await Promise.all([
    Store.find().lean(),
    LabelSerial.find().lean(),
  ]);
  for (const labelSerial of labelSerialsDocs) {
    labelSerialsMap.set(String(labelSerial.store), labelSerial.serialNumber);
  }

  if (req.user.role && req.user.role.name === 'PGP_administrator') {
    if (serialNumberBelgrade) {
      const store = stores.find((el) => el.name === 'Belgrade');
      labelSerialsMap.set(String(store._id), serialNumberBelgrade);
    }
    if (serialNumberBudapest) {
      const store = stores.find((el) => el.name === 'Budapest');
      labelSerialsMap.set(String(store._id), serialNumberBudapest);
    }
    if (serialNumberMontenegro) {
      const store = stores.find((el) => el.name === 'Porto Montenegro');
      labelSerialsMap.set(String(store._id), serialNumberMontenegro);
    }
  }

  const results = [];
  for await (const doc of cursorProducts) {
    try {
      const labelText = getProductLabelText(doc, labelSerialsMap);
      const declaration = getProductDeclaration(doc);
      const qrCodeLink = await getProductQRDataURL(doc, true);
      doc.boutiques.serialNumbers.labelText = labelText;
      doc.boutiques.serialNumbers.declaration = declaration;
      doc.boutiques.serialNumbers.qrCodeLink = qrCodeLink;
      results.push(doc);
    } catch (err) {
      results.push(doc);
    }
  }

  if (soonInStock === 'true') {
    const cursorSoonInStock = SoonInStock.aggregate(aggregatePipelineSoonInStock);
    for await (let doc of cursorSoonInStock) {
      try {
        doc = transformSoonInStockToProductDoc(doc);
        const labelText = getProductLabelText(doc, labelSerialsMap);
        const declaration = getProductDeclaration(doc);
        const qrCodeLink = await getProductQRDataURL(doc, true);
        doc.boutiques.serialNumbers.labelText = labelText;
        doc.boutiques.serialNumbers.declaration = declaration;
        doc.boutiques.serialNumbers.qrCodeLink = qrCodeLink;
        results.push(doc);
      } catch (err) {
        results.push(doc);
      }
    }
  }

  return res.status(200).send({
    message: 'Products with pgp references successfully fetched',
    results,
  });
};

/**
 * @api {post} /product/labels Print labels pdf
 * @apiVersion 1.0.0
 * @apiName Print labels pdf
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 *
 * @apiParam (body) {Boolean} [preview=false] Preview labels pdf
 * @apiParam (body) {String} [serialNumberBelgrade] Serial number for all labels for products from Belgrade
 * @apiParam (body) {String} [serialNumberBudapest] Serial number for all labels for products from Budapest
 * @apiParam (body) {String} [serialNumberMontenegro] Serial number for all labels for products from Montenegro
 * @apiParam (body) {Object[]} products Array with objects that contain information for each label
 * @apiParam (body) {Number} products.index Label index, position, starts from 0
 * @apiParam (body) {String="front", "back"} products.side Label side
 * @apiParam (body) {String} products.pgpReference Pgp reference
 * @apiParam (body) {String} [products.text] Edited label or declaration text
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
{
  "message": "Product labels PDF successfully created",
  "results": {
    "type": "products",
    "mimetype": "application/pdf",
    "fileName": "preview-product-labels-613f1391af2dbbe85c4783e1.pdf",
    "uploadedBy": "613f1391af2dbbe85c4783e1"
  }
}
 * @apiUse MissingParamsError
 * @apiUse InvalidValue
 * @apiUse NotFound
 */
module.exports.printProductLabels = async (req, res) => {
  let { serialNumberBelgrade, serialNumberBudapest, serialNumberMontenegro, products, preview } = req.body;
  const { _id: createdBy } = req.user;

  // check serial number
  const labelSerialsMap = new Map();
  const [stores, labelSerialsDocs] = await Promise.all([
    Store.find().lean(),
    LabelSerial.find().lean(),
  ]);
  for (const labelSerial of labelSerialsDocs) {
    labelSerialsMap.set(String(labelSerial.store), labelSerial.serialNumber);
  }

  if (req.user.role && req.user.role.name === 'PGP_administrator') {
    const toExecute = [];
    if (serialNumberBelgrade) {
      const store = stores.find((el) => el.name === 'Belgrade');
      if (store) toExecute.push(LabelSerial.updateOne({ store: store._id }, { serialNumber: serialNumberBelgrade, createdBy, store }, { upsert: true }));
      labelSerialsMap.set(String(store._id), serialNumberBelgrade);
    }
    if (serialNumberBudapest) {
      const store = stores.find((el) => el.name === 'Budapest');
      if (store) toExecute.push(LabelSerial.updateOne({ store: store._id }, { serialNumber: serialNumberBudapest, createdBy, store }, { upsert: true }));
      labelSerialsMap.set(String(store._id), serialNumberBudapest);
    }
    if (serialNumberMontenegro) {
      const store = stores.find((el) => el.name === 'Porto Montenegro');
      if (store) toExecute.push(LabelSerial.updateOne({ store: store._id }, { serialNumber: serialNumberMontenegro, createdBy, store }, { upsert: true }));
      labelSerialsMap.set(String(store._id), serialNumberMontenegro);
    }
    if (toExecute.length) {
      await Promise.all(toExecute);
    }
  }

  if (!products || !Array.isArray(products)) throw new Error(error.INVALID_VALUE);

  const allowedSides = ['front', 'back'];
  const indexSet = new Set();
  const sortedProducts = [];
  const productsMap = new Map();

  for (const product of products) {
    const { index, side, text, pgpReference } = product;

    if (!pgpReference || index === undefined || !side) throw new Error(error.MISSING_PARAMETERS);
    if (!allowedSides.includes(side)) throw new Error(error.INVALID_VALUE);

    // Check if index has valid value and if its unique
    if (typeof index !== 'number' || index < 0 || indexSet.has(index)) throw new Error(error.INVALID_VALUE);
    indexSet.add(index);

    // Create map with pgp key and aray with objects as value
    const producstArr = productsMap.get(pgpReference);
    if (producstArr) producstArr.push(product);
    else productsMap.set(pgpReference, [product]);

    sortedProducts[index] = product;
  }

  // Find all products
  const pgpArray = Array.from(productsMap.keys());
  const [soonInStockDocs, productsDocs] = await Promise.all([
    SoonInStock.find({ pgpReference: { $in: pgpArray } }).populate('product').lean(),
    Product.aggregate([
      { $unwind: '$boutiques' },
      { $unwind: '$boutiques.serialNumbers' },
      { $match: { 'boutiques.serialNumbers.pgpReference': { $in: pgpArray} } }
    ])
  ]);
  const pgpToFind = new Set(pgpArray);
  for (const doc of productsDocs) {
    const pgp = doc.boutiques.serialNumbers.pgpReference;
    pgpToFind.delete(pgp);
    const producstArr = productsMap.get(pgp);
    for (const product of producstArr) {
      product.document = doc;
    }
  }
  for (let doc of soonInStockDocs) {
    doc = transformSoonInStockToProductDoc(doc);
    const pgp = doc.boutiques.serialNumbers.pgpReference;
    pgpToFind.delete(pgp);
    const producstArr = productsMap.get(pgp);
    for (const product of producstArr) {
      product.document = doc;
    }
  }

  // Check if all documents are found
  if (pgpToFind.size > 0) throw new Error(error.NOT_FOUND);

  // Prepare data structure for pdf kit
  const pages = [[]];
  let perPage = 79;
  for (let i = 0; i < sortedProducts.length; i += 1) {
    if (!sortedProducts[i]) continue;

    if (sortedProducts[i].index > perPage) {
      pages.push([]);
      perPage += 80;
    }
    pages[pages.length - 1].push(sortedProducts[i]);
  }

  // Set print to false when developing
  const print = true;

  // Delete preview
  if (environments.NODE_ENV !== 'test') {
    if (!preview && print) deleteFile(`preview-product-labels-${createdBy}.pdf`);
  }

  // Create PDF
  const fileName = preview
    ? `preview-product-labels-${createdBy}.pdf`
    : `product-labels-${uuidv4()}.pdf`;

  const url = await labelsPDF({
    pages,
    preview,
    print,
    fileName,
    labelSerialsMap,
  });

  const document = {
    type: 'products',
    mimetype: 'application/pdf',
    fileName,
    url,
    uploadedBy: createdBy,
  };

  return res.status(200).send({
    message: 'Product labels PDF successfully created',
    results: document,
  });
};

/**
 * @api {get} /product/labels/serial-number Get saved serial number for product labels
 * @apiVersion 1.0.0
 * @apiName Get saved serial number for product labels
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - PGP_administrator
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
{
  "message": "Label serial number successfully fetched",
  "results": [
    {
      "_id": "61ae78a00ec2a669fe8b6c1c",
      "store": {
        "_id": "5f3a4225ffe375404f72fb06",
        "name": "Belgrade",
        "location": "Belgrade",
        "logo": "https://keanu.info",
        "__v": 0,
        "createdAt": "2021-03-01T08:58:44.019Z",
        "updatedAt": "2021-03-01T08:58:44.019Z",
        "vatPercent": 20
      },
      "serialNumber": "A120",
      "createdBy": "613f1391af2dbbe85c4783e1"
      "createdAt": "2021-12-06T14:05:58.115Z",
      "updatedAt": "2021-12-06T14:05:58.115Z",
      "__v": 0,
    },
    {
      "_id": "61ae78a00ec2a669fe8b6c1d",
      "store": {
        "_id": "5f3a4225ffe375404f72fb07",
        "name": "Budapest",
        "location": "Budapest",
        "logo": "http://hiram.info",
        "__v": 0,
        "createdAt": "2021-03-01T08:58:44.020Z",
        "updatedAt": "2021-03-01T08:58:44.020Z",
        "vatPercent": 27
      },
      "serialNumber": "B120",
      "createdBy": "613f1391af2dbbe85c4783e1"
      "createdAt": "2021-12-06T14:05:58.115Z",
      "updatedAt": "2021-12-06T14:05:58.115Z",
      "__v": 0,
    },
    {
      "_id": "61ae78a00ec2a669fe8b6c1f",
      "store": {
        "_id": "5f3a4225ffe375404f72fb08",
        "name": "Porto Montenegro",
        "location": "Tivat",
        "logo": "https://america.com",
        "__v": 0,
        "createdAt": "2021-03-01T08:58:44.020Z",
        "updatedAt": "2021-03-01T08:58:44.020Z",
        "vatPercent": 21
      },
      "serialNumber": "C120",
      "createdBy": "613f1391af2dbbe85c4783e1"
      "createdAt": "2021-12-06T14:05:58.115Z",
      "updatedAt": "2021-12-06T14:05:58.115Z",
      "__v": 0,
    }
  ]
}
 * @apiUse CredentialsError
 * @apiUse UnauthorizedError
 * @apiUse NotFound
 */
module.exports.getLabelSerial = async (req, res) => {
  const { user } = req;

  if (!user.role || user.role.name !== 'PGP_administrator') throw new Error(error.UNAUTHORIZED_ERROR);

  const results = await LabelSerial.find().populate('store').lean();

  if (!results) throw new Error(error.NOT_FOUND);

  res.status(200).send({
    message: 'Label serial number successfully fetched',
    results,
  });
};

/**
 * @api {post} /product/labels/serial-number Save serial number for product labels
 * @apiVersion 1.0.0
 * @apiName Save serial number for product labels
 * @apiGroup Product
 * @apiDescription ## Allowed roles:
 * - PGP_administrator
 *
 * @apiParam (body) {String} [serialNumberBelgrade] Serial number for all labels for products from Belgrade
 * @apiParam (body) {String} [serialNumberBudapest] Serial number for all labels for products from Budapest
 * @apiParam (body) {String} [serialNumberMontenegro] Serial number for all labels for products from Montenegro
 *
 * @apiSuccessExample Success-Response:
 HTTP/1.1 200 OK
{
  "message": "Label serial number successfully saved",
  "results": [
    {
      "_id": "61ae78a00ec2a669fe8b6c1c",
      "store": {
        "_id": "5f3a4225ffe375404f72fb06",
        "name": "Belgrade",
        "location": "Belgrade",
        "logo": "https://keanu.info",
        "__v": 0,
        "createdAt": "2021-03-01T08:58:44.019Z",
        "updatedAt": "2021-03-01T08:58:44.019Z",
        "vatPercent": 20
      },
      "serialNumber": "A120",
      "createdBy": "613f1391af2dbbe85c4783e1"
      "createdAt": "2021-12-06T14:05:58.115Z",
      "updatedAt": "2021-12-06T14:05:58.115Z",
      "__v": 0,
    },
    {
      "_id": "61ae78a00ec2a669fe8b6c1d",
      "store": {
        "_id": "5f3a4225ffe375404f72fb07",
        "name": "Budapest",
        "location": "Budapest",
        "logo": "http://hiram.info",
        "__v": 0,
        "createdAt": "2021-03-01T08:58:44.020Z",
        "updatedAt": "2021-03-01T08:58:44.020Z",
        "vatPercent": 27
      },
      "serialNumber": "B120",
      "createdBy": "613f1391af2dbbe85c4783e1"
      "createdAt": "2021-12-06T14:05:58.115Z",
      "updatedAt": "2021-12-06T14:05:58.115Z",
      "__v": 0,
    },
    {
      "_id": "61ae78a00ec2a669fe8b6c1f",
      "store": {
        "_id": "5f3a4225ffe375404f72fb08",
        "name": "Porto Montenegro",
        "location": "Tivat",
        "logo": "https://america.com",
        "__v": 0,
        "createdAt": "2021-03-01T08:58:44.020Z",
        "updatedAt": "2021-03-01T08:58:44.020Z",
        "vatPercent": 21
      },
      "serialNumber": "C120",
      "createdBy": "613f1391af2dbbe85c4783e1"
      "createdAt": "2021-12-06T14:05:58.115Z",
      "updatedAt": "2021-12-06T14:05:58.115Z",
      "__v": 0,
    }
  ]
}
 * @apiUse CredentialsError
 * @apiUse UnauthorizedError
 * @apiUse MissingParamsError
 * @apiUse NotFound
 */
module.exports.saveLabelSerial = async (req, res) => {
  const { user } = req;
  const { _id: createdBy } = user;
  const { serialNumberBelgrade, serialNumberBudapest, serialNumberMontenegro } = req.body;

  if (!user.role || user.role.name !== 'PGP_administrator') throw new Error(error.UNAUTHORIZED_ERROR);
  if (!serialNumberBelgrade && !serialNumberBudapest && !serialNumberMontenegro) throw new Error(error.MISSING_PARAMETERS);

  const toExecute = [];
  const stores = await Store.find();
  if (serialNumberBelgrade) {
    const store = stores.find((el) => el.name === 'Belgrade');
    if (store) toExecute.push(LabelSerial.updateOne({ store: store._id }, { serialNumber: serialNumberBelgrade, createdBy, store }, { upsert: true }));
  }
  if (serialNumberBudapest) {
    const store = stores.find((el) => el.name === 'Budapest');
    if (store) toExecute.push(LabelSerial.updateOne({ store: store._id }, { serialNumber: serialNumberBudapest, createdBy, store }, { upsert: true }));
  }
  if (serialNumberMontenegro) {
    const store = stores.find((el) => el.name === 'Porto Montenegro');
    if (store) toExecute.push(LabelSerial.updateOne({ store: store._id }, { serialNumber: serialNumberMontenegro, createdBy, store }, { upsert: true }));
  }

  await Promise.all(toExecute);

  const results = await LabelSerial.find().populate('store').lean();

  res.status(200).send({
    message: 'Label serial number successfully saved',
    results,
  });
};
