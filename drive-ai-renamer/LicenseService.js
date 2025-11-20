/**
 * Service to handle license key validation and storage.
 */
const LicenseService = {
  /**
   * Validates a license key.
   * For now, this is a mock implementation.
   * In production, you would call your Gumroad/LemonSqueezy API here.
   * 
   * @param {string} key The license key to validate.
   * @returns {boolean} True if valid, false otherwise.
   */
  validateLicenseKey: function(key) {
    if (!key || key.trim() === "") return false;
    
    // Mock validation: any key starting with "PRO-" is valid
    // TODO: Replace with actual API call to your payment provider
    return key.trim().toUpperCase().startsWith("PRO-");
  },

  /**
   * Saves the license key to user properties.
   * @param {string} key The license key.
   */
  saveLicenseKey: function(key) {
    PropertiesService.getUserProperties().setProperty("LICENSE_KEY", key.trim());
  },

  /**
   * Gets the stored license key.
   * @returns {string} The license key or null.
   */
  getLicenseKey: function() {
    return PropertiesService.getUserProperties().getProperty("LICENSE_KEY");
  },

  /**
   * Checks if the user has a valid license.
   * @returns {boolean}
   */
  hasValidLicense: function() {
    const key = this.getLicenseKey();
    return this.validateLicenseKey(key);
  },

  /**
   * Returns the file limit based on license status.
   * @returns {number} Max number of files allowed.
   */
  getFileLimit: function() {
    return this.hasValidLicense() ? Infinity : 10;
  }
};
