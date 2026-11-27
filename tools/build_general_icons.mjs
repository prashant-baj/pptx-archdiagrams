// Build the bundled "general" icon pack: black glyphs on transparent 256 px PNGs.
// archdiagrams tints them to any colour at render time.
// Sources: Tabler Icons (MIT) and Material Design Icons (Apache 2.0) via react-icons.
//
//   cd tools && npm install && node build_general_icons.mjs
//
// To add an icon: add a line to ICONS (friendly-name: react-icons component name), run the script, commit the PNG.
import React from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import sharp from 'sharp';
import * as tb from 'react-icons/tb';
import * as md from 'react-icons/md';
import { mkdirSync, writeFileSync } from 'fs';

const OUT = new URL('../src/archdiagrams/icons/general/', import.meta.url).pathname;

export const ICONS = {
  // people
  user: 'TbUser', users: 'TbUsers', admin: 'TbUserCog', 'user-secure': 'TbUserShield', team: 'MdGroups',
  engineer: 'MdEngineering', support: 'MdSupportAgent', headset: 'TbHeadset',
  // devices and clients
  mobile: 'TbDeviceMobile', laptop: 'TbDeviceLaptop', desktop: 'TbDeviceDesktop', tablet: 'TbDeviceTablet',
  browser: 'TbWorldWww', terminal: 'TbTerminal2', camera: 'TbCamera', printer: 'TbPrinter', phone: 'TbPhone',
  // compute, network, platform
  server: 'TbServer', 'server-rack': 'TbServer2', 'server-dns': 'MdDns', cloud: 'TbCloud', 'cloud-compute': 'TbCloudComputing',
  'cloud-lock': 'TbCloudLock', 'cloud-upload': 'TbCloudUpload', 'cloud-download': 'TbCloudDownload',
  container: 'TbContainer', box: 'TbBox', boxes: 'TbBoxMultiple', packages: 'TbPackages', function: 'TbFunction',
  router: 'TbRouter', gateway: 'TbRouter', network: 'MdLan', 'load-balancer': 'TbLoadBalancer', firewall: 'TbWall',
  internet: 'TbWorld', wifi: 'TbWifi', antenna: 'TbAntenna', bluetooth: 'TbBluetooth', tower: 'TbBuildingBroadcastTower',
  satellite: 'TbSatellite', hub: 'TbTopologyStar3', 'hub-alt': 'MdHub', webhook: 'TbWebhook', link: 'TbLink',
  // data and integration
  database: 'TbDatabase', 'database-export': 'TbDatabaseExport', storage: 'MdStorage', table: 'TbTable',
  queue: 'MdQueue', stream: 'TbArrowsShuffle', exchange: 'TbArrowsExchange', cache: 'TbBolt', api: 'TbApi',
  layers: 'TbStack2', archive: 'TbArchive', backup: 'TbDeviceFloppy', inbox: 'TbInbox', 'data-tree': 'TbBinaryTree',
  integration: 'TbLayersIntersect', sync: 'TbRefresh', file: 'TbFile', document: 'TbFileText', folder: 'TbFolder',
  spreadsheet: 'TbFileSpreadsheet', report: 'TbReport', logs: 'TbList',
  // security
  lock: 'TbLock', key: 'TbKey', shield: 'TbShield', 'shield-check': 'TbShieldCheck', security: 'MdSecurity',
  fingerprint: 'TbFingerprint', certificate: 'TbCertificate', identity: 'TbUserShield',
  // apps, ops and analytics
  dashboard: 'TbLayoutDashboard', analytics: 'TbDeviceDesktopAnalytics', 'chart-bar': 'TbChartBar', 'chart-line': 'TbChartLine',
  monitor: 'TbHeartRateMonitor', observe: 'TbEye', alert: 'TbBell', warning: 'TbAlertTriangle', check: 'TbCheck',
  search: 'TbSearch', settings: 'TbSettings', code: 'TbCode', git: 'TbBrandGit', branch: 'TbGitBranch', bug: 'TbBug',
  test: 'TbTestPipe', deploy: 'TbRocket', update: 'MdSystemUpdate', workflow: 'TbSitemap', hierarchy: 'TbHierarchy',
  schedule: 'TbClock', calendar: 'TbCalendar', mail: 'TbMail', chat: 'TbMessage', ticket: 'TbTicket', map: 'TbMap',
  calculator: 'TbCalculator', presentation: 'TbPresentation',
  // AI
  ai: 'TbBrain', robot: 'TbRobot', chip: 'TbCpu', memory: 'MdMemory',
  // business
  payment: 'TbCreditCard', payments: 'MdPayments', cart: 'TbShoppingCart', store: 'TbBuildingStore', bank: 'TbBuildingBank',
  invoice: 'TbReceipt', briefcase: 'TbBriefcase', scale: 'TbScale', handshake: 'TbHeartHandshake', school: 'TbSchool',
  building: 'TbBuilding', home: 'TbHome', hospital: 'TbBuildingHospital', warehouse: 'TbBuildingWarehouse',
  factory: 'TbBuildingFactory2', 'factory-alt': 'MdFactory', truck: 'TbTruck', car: 'TbCar',
  // energy and industrial (OT / IoT)
  power: 'MdPower', electrical: 'MdElectricalServices', bolt: 'TbBolt', battery: 'TbBattery', ups: 'TbBatteryCharging',
  plug: 'TbPlug', pdu: 'TbPlug', meter: 'TbGauge', gauge: 'TbGauge', breaker: 'TbCircuitSwitchClosed',
  motor: 'TbEngine', drive: 'MdSettingsInputComponent', plc: 'MdPrecisionManufacturing', sensor: 'MdSensors',
  temperature: 'TbTemperature', cooling: 'TbPropeller', 'ev-charger': 'TbChargingPile', 'ev-station': 'MdEvStation',
  solar: 'TbSolarPanel', wind: 'TbWindmill', water: 'TbDroplet', gas: 'TbFlame', leaf: 'TbLeaf', grid: 'TbBuildingFactory2',
};

mkdirSync(OUT, { recursive: true });
const lib = (n) => (n.startsWith('Tb') ? tb : md)[n];
for (const [name, comp] of Object.entries(ICONS)) {
  const C = lib(comp);
  if (!C) throw new Error(`unknown component ${comp}`);
  const svg = renderToStaticMarkup(React.createElement(C, { size: 256, color: '#000000' }));
  await sharp(Buffer.from(svg)).resize(256, 256).png().toFile(`${OUT}${name}.png`);
}
writeFileSync(`${OUT}SOURCES.json`, JSON.stringify(ICONS, null, 1));
console.log(`${Object.keys(ICONS).length} icons -> ${OUT}`);
