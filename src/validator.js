export const MISSING_KEY_FILENAME = 'Zipclex config missing property filename';

export const INVALID_TYPE_FILENAME = 'Zipclex filename can only be of type string';

export const INVALID_TYPE_SHEET = 'Zipcelx sheet data is not of type array';

export const INVALID_TYPE_SHEET_DATA = 'Zipclex sheet data childs is not of type array';

const childValidator = (array) => {
  return array.every((item) => Array.isArray(item));
};

export default (config) => {
  if (!config.filename) {
    console.error(MISSING_KEY_FILENAME);
    return false;
  }

  if (typeof config.filename !== 'string') {
    console.error(INVALID_TYPE_FILENAME);
    return false;
  }

  if (!Array.isArray(config.sheet.data)) {
    console.error(INVALID_TYPE_SHEET);
    return false;
  }

  if (!childValidator(config.sheet.data)) {
    console.error(INVALID_TYPE_SHEET_DATA);
    return false;
  }

  return true;
};
