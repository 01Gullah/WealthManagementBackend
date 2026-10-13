/**
 * WM_00_Config.gs
 * Apex Wealth - global constants.
 * Used by: WM_01_Web_App_Handler.gs (PROJECT_NAME, VERSION).
 * Balances start at 0 (no demo funds). Rules live in the Google Sheet, not here.
 */

var CONFIG = {
  PROJECT_NAME: 'Apex Wealth',
  VERSION: '2.0.0',
  SUPPORTED_ASSETS: ['BTC', 'ETH', 'USDT', 'SOL']
};
