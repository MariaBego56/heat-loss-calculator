import React, { useState } from 'react';
import { Download, Plus, Trash2, Home, User, Calendar, AlertCircle } from 'lucide-react';
import * as XLSX from 'xlsx';

const ROOM_TYPE_DATA = {
  'Living Room': { ach: 1.0, defaultTemp: 21 },
  'Second Living Room': { ach: 1.0, defaultTemp: 21 },
  'Kitchen': { ach: 2.0, defaultTemp: 18 },
  'Bedroom 1': { ach: 0.5, defaultTemp: 18 },
  'Bedroom 2': { ach: 0.5, defaultTemp: 18 },
  'Bedroom 3': { ach: 0.5, defaultTemp: 18 },
  'Bedroom 4': { ach: 0.5, defaultTemp: 18 },
  'Hallway': { ach: 1.5, defaultTemp: 18 },
  'Middle Hallway': { ach: 1.5, defaultTemp: 18 },
  'Upper Hallway': { ach: 1.5, defaultTemp: 18 },
  'Bathroom 1': { ach: 3.0, defaultTemp: 22 },
  'Bathroom 2': { ach: 3.0, defaultTemp: 22 },
  'Dining Room': { ach: 1.0, defaultTemp: 21 },
  'Study': { ach: 1.0, defaultTemp: 21 },
  'Utility Room': { ach: 2.0, defaultTemp: 16 },
  'Loo': { ach: 2.0, defaultTemp: 18 },
  'Laundry Room': { ach: 2.0, defaultTemp: 16 },
  'Pantry': { ach: 0.5, defaultTemp: 12 },
  'Basement': { ach: 0.5, defaultTemp: 18 },
  'En-Suite': { ach: 3.0, defaultTemp: 22 },
  'Shower Room': { ach: 3.0, defaultTemp: 22 }
};
const ROOM_TYPES = Object.keys(ROOM_TYPE_DATA);

const WALL_TYPES = {
  'No External Wall': 0,
  'Solid Wall (Uninsulated)': 2.1,
  'Solid Wall (Insulated)': 0.35,
  'Cavity Wall (Uninsulated)': 1.5,
  'Cavity Wall (Insulated)': 0.55,
  'Timber Frame': 0.3
};

const WINDOW_TYPES = {
  'Single Glazed - Wood': 4.8,
  'Single Glazed - Aluminium': 5.7,
  'Single Glazed - PVC': 4.8,
  'Double Glazed - Wood': 2.8,
  'Double Glazed - Aluminium': 3.4,
  'Double Glazed - PVC': 2.8,
  'Double Glazed Low-E - Wood': 1.8,
  'Double Glazed Low-E - PVC': 1.8,
  'Triple Glazed - Wood': 2.1,
  'Triple Glazed - Aluminium': 2.6,
  'Triple Glazed - PVC': 2.1,
  'No Window': 0
};

const DOOR_TYPES = {
  'Solid Wood Door': 3.0,
  'Glazed Wood Door (Single)': 5.7,
  'Glazed Wood Door (Double)': 3.4,
  'Metal Door (Single Glazed)': 5.7,
  'Metal Door (Double Glazed)': 3.4,
  'Composite Door': 1.8,
  'No Door': 0
};

const RADIATOR_TYPES = [
  'Single Panel Type 11 (K1)',
  'Single Convector Type 21 (K2)', 
  'Double Panel Type 22 (K3)',
  'Compact Type 11',
  'Compact Type 21',
  'Compact Type 22',
  'Towel Rail - Chrome',
  'Towel Rail - White'
];

// Radiator outputs at ΔT50 in Watts - organized by type, then height, then length
// Based on standard UK radiator manufacturer data
const RADIATOR_OUTPUTS = {
  'Single Panel Type 11 (K1)': {
    300: { 400: 298, 500: 372, 600: 447, 700: 521, 800: 596, 900: 670, 1000: 745, 1100: 819, 1200: 894, 1400: 1043, 1600: 1192, 1800: 1341, 2000: 1490 },
    400: { 400: 397, 500: 496, 600: 596, 700: 695, 800: 794, 900: 894, 1000: 993, 1100: 1092, 1200: 1192, 1400: 1390, 1600: 1589, 1800: 1788, 2000: 1986 },
    450: { 400: 447, 500: 558, 600: 670, 700: 782, 800: 894, 900: 1005, 1000: 1117, 1100: 1229, 1200: 1341, 1400: 1564, 1600: 1788, 1800: 2011, 2000: 2234 },
    500: { 400: 496, 500: 620, 600: 745, 700: 869, 800: 993, 900: 1117, 1000: 1241, 1100: 1365, 1200: 1490, 1400: 1738, 1600: 1986, 1800: 2234, 2000: 2482 },
    600: { 400: 596, 500: 745, 600: 894, 700: 1043, 800: 1192, 900: 1341, 1000: 1490, 1100: 1639, 1200: 1788, 1400: 2086, 1600: 2384, 1800: 2682, 2000: 2980 },
    700: { 400: 695, 500: 869, 600: 1043, 700: 1217, 800: 1390, 900: 1564, 1000: 1738, 1100: 1912, 1200: 2086, 1400: 2433, 1600: 2781, 1800: 3128, 2000: 3476 }
  },
  'Single Convector Type 21 (K2)': {
    300: { 400: 462, 500: 578, 600: 693, 700: 809, 800: 924, 900: 1040, 1000: 1155, 1100: 1271, 1200: 1386, 1400: 1617, 1600: 1848, 1800: 2079, 2000: 2310 },
    400: { 400: 616, 500: 770, 600: 924, 700: 1078, 800: 1232, 900: 1386, 1000: 1540, 1100: 1694, 1200: 1848, 1400: 2156, 1600: 2464, 1800: 2772, 2000: 3080 },
    450: { 400: 693, 500: 866, 600: 1040, 700: 1213, 800: 1386, 900: 1559, 1000: 1733, 1100: 1906, 1200: 2079, 1400: 2426, 1600: 2772, 1800: 3119, 2000: 3465 },
    500: { 400: 770, 500: 963, 600: 1155, 700: 1348, 800: 1540, 900: 1733, 1000: 1925, 1100: 2118, 1200: 2310, 1400: 2695, 1600: 3080, 1800: 3465, 2000: 3850 },
    600: { 400: 924, 500: 1155, 600: 1386, 700: 1617, 800: 1848, 900: 2079, 1000: 2310, 1100: 2541, 1200: 2772, 1400: 3234, 1600: 3696, 1800: 4158, 2000: 4620 },
    700: { 400: 1078, 500: 1348, 600: 1617, 700: 1887, 800: 2156, 900: 2426, 1000: 2695, 1100: 2965, 1200: 3234, 1400: 3773, 1600: 4312, 1800: 4851, 2000: 5390 }
  },
  'Double Panel Type 22 (K3)': {
    300: { 400: 522, 500: 652, 600: 783, 700: 913, 800: 1044, 900: 1174, 1000: 1305, 1100: 1435, 1200: 1566, 1400: 1827, 1600: 2088, 1800: 2349, 2000: 2610 },
    400: { 400: 696, 500: 870, 600: 1044, 700: 1218, 800: 1392, 900: 1566, 1000: 1740, 1100: 1914, 1200: 2088, 1400: 2436, 1600: 2784, 1800: 3132, 2000: 3480 },
    450: { 400: 783, 500: 979, 600: 1174, 700: 1370, 800: 1566, 900: 1762, 1000: 1958, 1100: 2153, 1200: 2349, 1400: 2741, 1600: 3132, 1800: 3524, 2000: 3915 },
    500: { 400: 870, 500: 1088, 600: 1305, 700: 1523, 800: 1740, 900: 1958, 1000: 2175, 1100: 2393, 1200: 2610, 1400: 3045, 1600: 3480, 1800: 3915, 2000: 4350 },
    600: { 400: 1044, 500: 1305, 600: 1566, 700: 1827, 800: 2088, 900: 2349, 1000: 2610, 1100: 2871, 1200: 3132, 1400: 3654, 1600: 4176, 1800: 4698, 2000: 5220 },
    700: { 400: 1218, 500: 1523, 600: 1827, 700: 2132, 800: 2436, 900: 2741, 1000: 3045, 1100: 3350, 1200: 3654, 1400: 4263, 1600: 4872, 1800: 5481, 2000: 6090 }
  },
  'Compact Type 11': {
    300: { 400: 298, 500: 372, 600: 447, 700: 521, 800: 596, 900: 670, 1000: 745, 1100: 819, 1200: 894, 1400: 1043, 1600: 1192, 1800: 1341, 2000: 1490 },
    400: { 400: 397, 500: 496, 600: 596, 700: 695, 800: 794, 900: 894, 1000: 993, 1100: 1092, 1200: 1192, 1400: 1390, 1600: 1589, 1800: 1788, 2000: 1986 },
    450: { 400: 447, 500: 558, 600: 670, 700: 782, 800: 894, 900: 1005, 1000: 1117, 1100: 1229, 1200: 1341, 1400: 1564, 1600: 1788, 1800: 2011, 2000: 2234 },
    500: { 400: 496, 500: 620, 600: 745, 700: 869, 800: 993, 900: 1117, 1000: 1241, 1100: 1365, 1200: 1490, 1400: 1738, 1600: 1986, 1800: 2234, 2000: 2482 },
    600: { 400: 596, 500: 745, 600: 894, 700: 1043, 800: 1192, 900: 1341, 1000: 1490, 1100: 1639, 1200: 1788, 1400: 2086, 1600: 2384, 1800: 2682, 2000: 2980 },
    700: { 400: 695, 500: 869, 600: 1043, 700: 1217, 800: 1390, 900: 1564, 1000: 1738, 1100: 1912, 1200: 2086, 1400: 2433, 1600: 2781, 1800: 3128, 2000: 3476 }
  },
  'Compact Type 21': {
    300: { 400: 462, 500: 578, 600: 693, 700: 809, 800: 924, 900: 1040, 1000: 1155, 1100: 1271, 1200: 1386, 1400: 1617, 1600: 1848, 1800: 2079, 2000: 2310 },
    400: { 400: 616, 500: 770, 600: 924, 700: 1078, 800: 1232, 900: 1386, 1000: 1540, 1100: 1694, 1200: 1848, 1400: 2156, 1600: 2464, 1800: 2772, 2000: 3080 },
    450: { 400: 693, 500: 866, 600: 1040, 700: 1213, 800: 1386, 900: 1559, 1000: 1733, 1100: 1906, 1200: 2079, 1400: 2426, 1600: 2772, 1800: 3119, 2000: 3465 },
    500: { 400: 770, 500: 963, 600: 1155, 700: 1348, 800: 1540, 900: 1733, 1000: 1925, 1100: 2118, 1200: 2310, 1400: 2695, 1600: 3080, 1800: 3465, 2000: 3850 },
    600: { 400: 924, 500: 1155, 600: 1386, 700: 1617, 800: 1848, 900: 2079, 1000: 2310, 1100: 2541, 1200: 2772, 1400: 3234, 1600: 3696, 1800: 4158, 2000: 4620 },
    700: { 400: 1078, 500: 1348, 600: 1617, 700: 1887, 800: 2156, 900: 2426, 1000: 2695, 1100: 2965, 1200: 3234, 1400: 3773, 1600: 4312, 1800: 4851, 2000: 5390 }
  },
  'Compact Type 22': {
    300: { 400: 522, 500: 652, 600: 783, 700: 913, 800: 1044, 900: 1174, 1000: 1305, 1100: 1435, 1200: 1566, 1400: 1827, 1600: 2088, 1800: 2349, 2000: 2610 },
    400: { 400: 696, 500: 870, 600: 1044, 700: 1218, 800: 1392, 900: 1566, 1000: 1740, 1100: 1914, 1200: 2088, 1400: 2436, 1600: 2784, 1800: 3132, 2000: 3480 },
    450: { 400: 783, 500: 979, 600: 1174, 700: 1370, 800: 1566, 900: 1762, 1000: 1958, 1100: 2153, 1200: 2349, 1400: 2741, 1600: 3132, 1800: 3524, 2000: 3915 },
    500: { 400: 870, 500: 1088, 600: 1305, 700: 1523, 800: 1740, 900: 1958, 1000: 2175, 1100: 2393, 1200: 2610, 1400: 3045, 1600: 3480, 1800: 3915, 2000: 4350 },
    600: { 400: 1044, 500: 1305, 600: 1566, 700: 1827, 800: 2088, 900: 2349, 1000: 2610, 1100: 2871, 1200: 3132, 1400: 3654, 1600: 4176, 1800: 4698, 2000: 5220 },
    700: { 400: 1218, 500: 1523, 600: 1827, 700: 2132, 800: 2436, 900: 2741, 1000: 3045, 1100: 3350, 1200: 3654, 1400: 4263, 1600: 4872, 1800: 5481, 2000: 6090 }
  },
  'Towel Rail - Chrome':{
  800: { 400: 200, 500: 250, 600: 300 },
  1000: { 400: 280, 500: 350, 600: 420 },
  1200: { 400: 350, 500: 440, 600: 530 },
  1500: { 400: 450, 500: 560, 600: 670 },
  1800: { 400: 550, 500: 690, 600: 830 }
},
'Towel Rail - White': {
  800: { 400: 220, 500: 275, 600: 330 },
  1000: { 400: 310, 500: 385, 600: 460 },
  1200: { 400: 385, 500: 480, 600: 575 },
  1500: { 400: 495, 500: 615, 600: 735 },
  1800: { 400: 600, 500: 750, 600: 900 }
}

};

// Building age for thermal bridging
const BUILDING_AGE = {
  'Pre-1965': 0.15,
  '1965-1982': 0.12,
  '1983-1995': 0.10,
  '1996-2006': 0.08,
  'Post-2006': 0.05
};

// Heating pattern for intermittent uplift
const HEATING_PATTERN = {
  'Continuous (24 hours)': 0,
  '16 hours per day': 0.10,
  '12 hours per day': 0.15,
  '9 hours per day': 0.20
};

// Function to lookup radiator output, with interpolation for non-standard sizes
function getRadiatorOutput(type, height, length) {
  if (!type || !height || !length) return null;
  
  const typeData = RADIATOR_OUTPUTS[type];
  if (!typeData) return null;
  
  // Get available heights and find closest
  const heights = Object.keys(typeData).map(Number).sort((a, b) => a - b);
  let closestHeight = heights.reduce((prev, curr) => 
    Math.abs(curr - height) < Math.abs(prev - height) ? curr : prev
  );
  
  const lengthData = typeData[closestHeight];
  if (!lengthData) return null;
  
  // Get available lengths and find closest
  const lengths = Object.keys(lengthData).map(Number).sort((a, b) => a - b);
  let closestLength = lengths.reduce((prev, curr) => 
    Math.abs(curr - length) < Math.abs(prev - length) ? curr : prev
  );
  
  // Get base output for closest size
  const baseOutput = lengthData[closestLength];
  
  // Apply linear interpolation/extrapolation for actual size
  const heightRatio = height / closestHeight;
  const lengthRatio = length / closestLength;
  
  return Math.round(baseOutput * heightRatio * lengthRatio);
}

const ROOM_TEMP_OPTIONS = ['Heated Room', 'Unheated Room', 'Outside'];

const WINDOW_CONDITIONS = {
  'Good Condition': { factor: 1.0, infiltration: 0 },
  'Minor Seal Wear': { factor: 1.1, infiltration: 0.5 },
  'Damaged Seals': { factor: 1.25, infiltration: 1.5 },
  'Failed Seals / Drafty': { factor: 1.4, infiltration: 3.0 }
};

// Floor types with U-values (W/m²K)
const FLOOR_TYPES = {
  'Heated Room Below': { uValue: 0, tempDiff: 0 },
  'Unheated Room Below': { uValue: 0.25, tempDiff: 10 },
  'Insulated Suspended Timber Floor': { uValue: 0.25, tempDiff: 'external' },
  'Uninsulated Suspended Timber Floor': { uValue: 0.7, tempDiff: 'external' },
  'Insulated Concrete Floor (Slab on Ground)': { uValue: 0.25, tempDiff: 'ground' },
  'Uninsulated Concrete Floor (Slab on Ground)': { uValue: 0.7, tempDiff: 'ground' },
  'Insulated Concrete Floor (Over Void/Garage)': { uValue: 0.25, tempDiff: 'external' },
  'Uninsulated Concrete Floor (Over Void/Garage)': { uValue: 1.0, tempDiff: 'external' }
};

// Ceiling/Roof types with U-values (W/m²K)
const CEILING_TYPES = {
  'Heated Room Above': { uValue: 0, tempDiff: 0 },
  'Unheated Room Above': { uValue: 0.16, tempDiff: 10 },
  'Insulated Loft (270mm+ insulation)': { uValue: 0.16, tempDiff: 'external' },
  'Partially Insulated Loft (100mm insulation)': { uValue: 0.3, tempDiff: 'external' },
  'Uninsulated Loft': { uValue: 2.3, tempDiff: 'external' },
  'Insulated Flat Roof': { uValue: 0.25, tempDiff: 'external' },
  'Uninsulated Flat Roof': { uValue: 1.5, tempDiff: 'external' },
  'Insulated Pitched Roof (Room in Roof)': { uValue: 0.2, tempDiff: 'external' },
  'Uninsulated Pitched Roof (Room in Roof)': { uValue: 2.0, tempDiff: 'external' }
};

function App() {
  const [propertyName, setPropertyName] = useState('');
  const [engineerName, setEngineerName] = useState('');
  const [assessmentDate, setAssessmentDate] = useState(new Date().toISOString().split('T')[0]);
  const [buildingAge, setBuildingAge] = useState('1983-1995');
  const [heatingPattern, setHeatingPattern] = useState('12 hours per day');
  const [rooms, setRooms] = useState([createEmptyRoom()]);

  function createEmptyRoom() {
    return {
      id: Date.now() + Math.random(),
      roomType: '',
      requiredTemp: '21',
      designTemp: '-3',
      height: '',
      length: '',
      width: '',
      volume: '',
      
      // Exposed walls (up to 3)
      wall1H: '',
      wall1W: '',
      wall1Area: '',
      wall2H: '',
      wall2W: '',
      wall2Area: '',
      wall3H: '',
      wall3W: '',
      wall3Area: '',
      totalWallArea: '',
      wallType: '',
      
      // Windows (up to 3)
      windowType: '',
      windowCondition: '',
      win1H: '',
      win1W: '',
      win1Area: '',
      win2H: '',
      win2W: '',
      win2Area: '',
      win3H: '',
      win3W: '',
      win3Area: '',
      totalWinArea: '',
      
      // External doors (up to 2)
      door1Type: '',
      door1H: '',
      door1W: '',
      door1Area: '',
      door2Type: '',
      door2H: '',
      door2W: '',
      door2Area: '',
      totalDoorArea: '',
      
      floorType: '',
      ceilingType: '',
      
      // Heat calculations
      fabricHeatLoss: '',
      ventilationHeatLoss: '',
      requiredHeat: '',
      
      // Mean water temperature
      flowTemp: '75',
      returnTemp: '65',
      mwt: '70',
      deltaT: '50',
      
      // Existing radiator
      existRadType: '',
      existRadH: '',
      existRadL: '',
      existRadOutputDT50: '',
      existRadActual: '',
      
      // Recommendation
      recommendation: 'Keep as is',
      recRadType: '',
      recRadH: '',
      recRadL: '',
      recRadSize: '',
      recRadOutputDT50: '',
      recRadActual: '',
      addRadType: '',
      addRadSize: '',
      addRadOutput: ''
    };
  }

  const addRoom = () => setRooms([...rooms, createEmptyRoom()]);
  
  const deleteRoom = (id) => {
    if (rooms.length > 1) {
      setRooms(rooms.filter(r => r.id !== id));
    }
  };

  const updateRoom = (id, field, value) => {
    setRooms(rooms.map(room => {
      if (room.id !== id) return room;
      
      const u = { ...room, [field]: value };
      
      //auto-set temperature when room type changes
      if (field === 'roomType' && ROOM_TYPE_DATA[value]) {
        u.requiredTemp = ROOM_TYPE_DATA[value].defaultTemp.toString();
}
      // Calculate room volume
      if (['height', 'length', 'width'].includes(field)) {
        const h = parseFloat(u.height) || 0;
        const l = parseFloat(u.length) || 0;
        const w = parseFloat(u.width) || 0;
        u.volume = h && l && w ? (h * l * w).toFixed(2) : '';
      }
      
      // Calculate wall areas
      if (field.includes('wall1')) {
        const h = parseFloat(u.wall1H) || 0;
        const w = parseFloat(u.wall1W) || 0;
        u.wall1Area = h && w ? (h * w).toFixed(2) : '';
      }
      if (field.includes('wall2')) {
        const h = parseFloat(u.wall2H) || 0;
        const w = parseFloat(u.wall2W) || 0;
        u.wall2Area = h && w ? (h * w).toFixed(2) : '';
      }
      if (field.includes('wall3')) {
        const h = parseFloat(u.wall3H) || 0;
        const w = parseFloat(u.wall3W) || 0;
        u.wall3Area = h && w ? (h * w).toFixed(2) : '';
      }
      
      // Calculate total wall area
      const a1 = parseFloat(u.wall1Area) || 0;
      const a2 = parseFloat(u.wall2Area) || 0;
      const a3 = parseFloat(u.wall3Area) || 0;
      u.totalWallArea = (a1 + a2 + a3).toFixed(2);
      
      // Calculate window areas
      if (field.includes('win1')) {
        const h = parseFloat(u.win1H) || 0;
        const w = parseFloat(u.win1W) || 0;
        u.win1Area = h && w ? (h * w).toFixed(2) : '';
      }
      if (field.includes('win2')) {
        const h = parseFloat(u.win2H) || 0;
        const w = parseFloat(u.win2W) || 0;
        u.win2Area = h && w ? (h * w).toFixed(2) : '';
      }
      if (field.includes('win3')) {
        const h = parseFloat(u.win3H) || 0;
        const w = parseFloat(u.win3W) || 0;
        u.win3Area = h && w ? (h * w).toFixed(2) : '';
      }
      
      // Calculate total window area
      const wa1 = parseFloat(u.win1Area) || 0;
      const wa2 = parseFloat(u.win2Area) || 0;
      const wa3 = parseFloat(u.win3Area) || 0;
      u.totalWinArea = (wa1 + wa2 + wa3).toFixed(2);
      
      // Calculate door areas
      if (field.includes('door1')) {
        const h = parseFloat(u.door1H) || 0;
        const w = parseFloat(u.door1W) || 0;
        u.door1Area = h && w ? (h * w).toFixed(2) : '';
      }
      if (field.includes('door2')) {
        const h = parseFloat(u.door2H) || 0;
        const w = parseFloat(u.door2W) || 0;
        u.door2Area = h && w ? (h * w).toFixed(2) : '';
      }
      
      // Calculate total door area
      const da1 = parseFloat(u.door1Area) || 0;
      const da2 = parseFloat(u.door2Area) || 0;
      u.totalDoorArea = (da1 + da2).toFixed(2);
      
      // Calculate heat loss - ALWAYS RECALCULATE
      const hasBasicData = u.requiredTemp && u.designTemp && u.length && u.width && u.height;
      
      if (hasBasicData) {
        const wallU = WALL_TYPES[u.wallType] || 1.5;
        const baseWinU = WINDOW_TYPES[u.windowType] || 0;
        const door1U = DOOR_TYPES[u.door1Type] || 0;
        const door2U = DOOR_TYPES[u.door2Type] || 0;
        
        // Apply window condition factor to U-value
        const windowConditionData = WINDOW_CONDITIONS[u.windowCondition] || { factor: 1.0, infiltration: 0 };
        const winU = baseWinU * windowConditionData.factor;
        
        const wallA = parseFloat(u.totalWallArea) || 0;
        const winA = parseFloat(u.totalWinArea) || 0;
        const door1A = parseFloat(u.door1Area) || 0;
        const door2A = parseFloat(u.door2Area) || 0;
        const totalDoorA = door1A + door2A;
        const netWall = Math.max(0, wallA - winA - totalDoorA);
        
        const tempDiff = (parseFloat(u.requiredTemp) || 21) - (parseFloat(u.designTemp) || -3);
        const roomTemp = parseFloat(u.requiredTemp) || 21;
        
        const wallLoss = netWall * wallU * tempDiff;
        const winLoss = winA * winU * tempDiff;
        const door1Loss = door1A * door1U * tempDiff;
        const door2Loss = door2A * door2U * tempDiff;
        const doorLoss = door1Loss + door2Loss;
        
        const floorArea = (parseFloat(u.length) || 0) * (parseFloat(u.width) || 0);
        
        // Calculate floor heat loss based on floor type
        let floorLoss = 0;
        const floorData = FLOOR_TYPES[u.floorType];
        if (floorData) {
          if (floorData.tempDiff === 0) {
            floorLoss = 0; // Heated room below
          } else if (floorData.tempDiff === 'external') {
            floorLoss = floorArea * floorData.uValue * tempDiff;
          } else if (floorData.tempDiff === 'ground') {
            // Ground temperature assumed ~10°C, so temp diff is room temp - 10
            const groundTempDiff = roomTemp - 10;
            floorLoss = floorArea * floorData.uValue * groundTempDiff;
          } else {
            // Fixed temp diff (e.g., unheated room = 10°C diff)
            floorLoss = floorArea * floorData.uValue * floorData.tempDiff;
          }
        }
        
        // Calculate ceiling heat loss based on ceiling type
        let ceilingLoss = 0;
        const ceilingData = CEILING_TYPES[u.ceilingType];
        if (ceilingData) {
          if (ceilingData.tempDiff === 0) {
            ceilingLoss = 0; // Heated room above
          } else if (ceilingData.tempDiff === 'external') {
            ceilingLoss = floorArea * ceilingData.uValue * tempDiff;
          } else {
            // Fixed temp diff (e.g., unheated room = 10°C diff)
            ceilingLoss = floorArea * ceilingData.uValue * ceilingData.tempDiff;
          }
        }
        
        u.fabricHeatLoss = (wallLoss + winLoss + doorLoss + floorLoss + ceilingLoss).toFixed(0);
        
        //Thermal bridging calculation
        const thermalBridgingFactor = BUILDING_AGE[buildingAge] || 0.10;
        const thermalBridgingLoss = parseFloat(u.fabricHeatLoss) * thermalBridgingFactor;
      
        const subtotal = parseFloat(u.fabricHeatLoss) + thermalBridgingLoss + parseFloat(u.ventilationHeatLoss);
        const intermittentFactor = HEATING_PATTERN[heatingPattern] || 0;
        const intermittentUplift = subtotal * intermittentFactor;
        u.requiredHeat = (subtotal + intermittentUplift).toFixed(0);
        
        // ventilation heat loss calculation with room-specific ACH
        const roomACH = ROOM_TYPE_DATA[u.roomType]?.ach || 1.0;
        const totalACH = roomACH + (windowConditionData.infiltration * (winA / 10));
        u.ventilationHeatLoss = (vol * totalACH * tempDiff * 0.33).toFixed(0);
        
        // Total required heat
        const totalLoss = parseFloat(u.fabricHeatLoss) + parseFloat(u.ventilationHeatLoss);
        u.requiredHeat = totalLoss.toFixed(0);
      } else {
        u.fabricHeatLoss = '';
        u.ventilationHeatLoss = '';
        u.requiredHeat = '';
      }
      
      // Always recalculate deltaT first (in case flow/return/room temp changed)
      const flow = parseFloat(u.flowTemp) || 75;
      const ret = parseFloat(u.returnTemp) || 65;
      const roomT = parseFloat(u.requiredTemp) || 21;
      u.mwt = ((flow + ret) / 2).toFixed(1);
      u.deltaT = (parseFloat(u.mwt) - roomT).toFixed(0);
      
      // Auto-lookup existing radiator output based on type and dimensions
      const dt = parseFloat(u.deltaT) || 50;
      if (u.existRadType && u.existRadH && u.existRadL) {
        const height = parseFloat(u.existRadH);
        const length = parseFloat(u.existRadL);
        const outputDT50 = getRadiatorOutput(u.existRadType, height, length);
        if (outputDT50) {
          u.existRadOutputDT50 = outputDT50.toString();
          const factor = Math.pow(dt / 50, 1.3);
          u.existRadActual = (outputDT50 * factor).toFixed(0);
        }
      } else {
        u.existRadOutputDT50 = '';
        u.existRadActual = '';
      }
      
      // Auto-lookup recommended radiator output based on type and dimensions
      if (u.recRadType && u.recRadH && u.recRadL) {
        const height = parseFloat(u.recRadH);
        const length = parseFloat(u.recRadL);
        const outputDT50 = getRadiatorOutput(u.recRadType, height, length);
        if (outputDT50) {
          u.recRadOutputDT50 = outputDT50.toString();
          const factor = Math.pow(dt / 50, 1.3);
          u.recRadActual = (outputDT50 * factor).toFixed(0);
        }
      } else {
        u.recRadOutputDT50 = '';
        u.recRadActual = '';
      }
      
      return u;
    }));
  };


  const generateExcel = () => {
    const wb = XLSX.utils.book_new();
    
    // Build the main assessment sheet with rooms as columns
    const data = [];
    
    // Header rows
    data.push(['HEAT LOSS ASSESSMENT REPORT', '', ...rooms.map(() => '')]);
    data.push(['Property:', propertyName, ...rooms.map(() => '')]);
    data.push(['Engineer:', engineerName, ...rooms.map(() => '')]);
    data.push(['Date:', assessmentDate, ...rooms.map(() => '')]);
    data.push(['System Flow/Return:', rooms[0] ? `${rooms[0].flowTemp}°C / ${rooms[0].returnTemp}°C` : '', ...rooms.map(() => '')]);
    data.push(['System Delta T:', rooms[0] ? `ΔT${rooms[0].deltaT}` : '', ...rooms.map(() => '')]);
    data.push([]); // Empty row
    
    // Column headers (Room numbers)
    data.push(['', ...rooms.map((r, i) => r.roomType || `Room ${i + 1}`)]);
    
    // Room data rows - each property as a row, rooms as columns
    data.push(['ROOM DETAILS', ...rooms.map(() => '')]);
    data.push(['Required Temperature (°C)', ...rooms.map(r => r.requiredTemp)]);
    data.push(['Design Temperature (°C)', ...rooms.map(r => r.designTemp)]);
    data.push(['Temperature Difference (°C)', ...rooms.map(r => r.requiredTemp && r.designTemp ? (parseFloat(r.requiredTemp) - parseFloat(r.designTemp)).toFixed(0) : '')]);
    data.push([]); // Empty row
    
    data.push(['ROOM DIMENSIONS', ...rooms.map(() => '')]);
    data.push(['Height (m)', ...rooms.map(r => r.height)]);
    data.push(['Length (m)', ...rooms.map(r => r.length)]);
    data.push(['Width (m)', ...rooms.map(r => r.width)]);
    data.push(['Volume (m³)', ...rooms.map(r => r.volume)]);
    data.push([]); // Empty row
    
    data.push(['EXTERNAL WALLS', ...rooms.map(() => '')]);
    data.push(['Wall Type', ...rooms.map(r => r.wallType)]);
    data.push(['Wall U-Value (W/m²K)', ...rooms.map(r => r.wallType ? WALL_TYPES[r.wallType] : '')]);
    data.push(['Wall 1 Area (m²)', ...rooms.map(r => r.wall1Area || '-')]);
    data.push(['Wall 2 Area (m²)', ...rooms.map(r => r.wall2Area || '-')]);
    data.push(['Wall 3 Area (m²)', ...rooms.map(r => r.wall3Area || '-')]);
    data.push(['Total Wall Area (m²)', ...rooms.map(r => r.totalWallArea)]);
    data.push([]); // Empty row
    
    data.push(['WINDOWS', ...rooms.map(() => '')]);
    data.push(['Window Type', ...rooms.map(r => r.windowType)]);
    data.push(['Window Condition', ...rooms.map(r => r.windowCondition)]);
    data.push(['Window U-Value (W/m²K)', ...rooms.map(r => r.windowType ? WINDOW_TYPES[r.windowType] : '')]);
    data.push(['Condition Factor', ...rooms.map(r => r.windowCondition && WINDOW_CONDITIONS[r.windowCondition] ? WINDOW_CONDITIONS[r.windowCondition].factor + 'x' : '')]);
    data.push(['Window 1 Area (m²)', ...rooms.map(r => r.win1Area || '-')]);
    data.push(['Window 2 Area (m²)', ...rooms.map(r => r.win2Area || '-')]);
    data.push(['Window 3 Area (m²)', ...rooms.map(r => r.win3Area || '-')]);
    data.push(['Total Window Area (m²)', ...rooms.map(r => r.totalWinArea)]);
    data.push([]); // Empty row
    
    data.push(['EXTERNAL DOOR', ...rooms.map(() => '')]);
    data.push(['Door 1 Type', ...rooms.map(r => r.door1Type || 'No Door')]);
    data.push(['Door 1 U-Value (W/m²K)', ...rooms.map(r => r.door1Type ? DOOR_TYPES[r.door1Type] : '-')]);
    data.push(['Door 1 Area (m²)', ...rooms.map(r => r.door1Area || '-')]);
    data.push(['Door 2 Type', ...rooms.map(r => r.door2Type || 'No Door')]);
    data.push(['Door 2 U-Value (W/m²K)', ...rooms.map(r => r.door2Type ? DOOR_TYPES[r.door2Type] : '-')]);
    data.push(['Door 2 Area (m²)', ...rooms.map(r => r.door2Area || '-')]);
    data.push(['Total Door Area (m²)', ...rooms.map(r => r.totalDoorArea || '-')]);
    data.push([]); // Empty row
    
    data.push(['FLOOR & CEILING', ...rooms.map(() => '')]);
    data.push(['Floor Type', ...rooms.map(r => r.floorType)]);
    data.push(['Floor U-Value (W/m²K)', ...rooms.map(r => r.floorType && FLOOR_TYPES[r.floorType] ? FLOOR_TYPES[r.floorType].uValue : '')]);
    data.push(['Ceiling/Roof Type', ...rooms.map(r => r.ceilingType)]);
    data.push(['Ceiling U-Value (W/m²K)', ...rooms.map(r => r.ceilingType && CEILING_TYPES[r.ceilingType] ? CEILING_TYPES[r.ceilingType].uValue : '')]);
    data.push([]); // Empty row
    
    data.push(['HEAT LOSS CALCULATION', ...rooms.map(() => '')]);
    data.push(['Fabric Heat Loss (W)', ...rooms.map(r => r.fabricHeatLoss)]);
    data.push(['Ventilation Heat Loss (W)', ...rooms.map(r => r.ventilationHeatLoss)]);
    data.push(['TOTAL Required Heat (W)', ...rooms.map(r => r.requiredHeat)]);
    data.push([]); // Empty row
    
    data.push(['EXISTING RADIATOR', ...rooms.map(() => '')]);
    data.push(['Radiator Type', ...rooms.map(r => r.existRadType)]);
    data.push(['Height (mm)', ...rooms.map(r => r.existRadH)]);
    data.push(['Length (mm)', ...rooms.map(r => r.existRadL)]);
    data.push(['Output @ ΔT50 (W)', ...rooms.map(r => r.existRadOutputDT50)]);
    data.push(['Actual Output (W)', ...rooms.map(r => r.existRadActual)]);
    data.push(['Status', ...rooms.map(r => {
      if (!r.existRadActual || !r.requiredHeat) return '';
      return parseFloat(r.existRadActual) >= parseFloat(r.requiredHeat) ? 'OK' : 'UNDERSIZED';
    })]);
    data.push(['Shortfall (W)', ...rooms.map(r => {
      if (!r.existRadActual || !r.requiredHeat) return '';
      const shortfall = parseFloat(r.requiredHeat) - parseFloat(r.existRadActual);
      return shortfall > 0 ? shortfall.toFixed(0) : '0';
    })]);
    data.push([]); // Empty row
    
    data.push(['RECOMMENDATION', ...rooms.map(() => '')]);
    data.push(['Action Required', ...rooms.map(r => r.recommendation)]);
    data.push(['Recommended Rad Type', ...rooms.map(r => r.recRadType || '-')]);
    data.push(['Recommended Height (mm)', ...rooms.map(r => r.recRadH || '-')]);
    data.push(['Recommended Length (mm)', ...rooms.map(r => r.recRadL || '-')]);
    data.push(['Recommended Output @ ΔT50 (W)', ...rooms.map(r => r.recRadOutputDT50 || '-')]);
    data.push(['Recommended Actual Output (W)', ...rooms.map(r => r.recRadActual || '-')]);
    data.push([]); // Empty row
    
    // Property totals row
    data.push(['PROPERTY TOTALS', ...rooms.map(() => '')]);
    const totalRequired = rooms.reduce((sum, r) => sum + (parseFloat(r.requiredHeat) || 0), 0);
    const totalExisting = rooms.reduce((sum, r) => sum + (parseFloat(r.existRadActual) || 0), 0);
    data.push(['Total Required Heat (W)', totalRequired.toFixed(0), ...rooms.slice(1).map(() => '')]);
    data.push(['Total Existing Output (W)', totalExisting.toFixed(0), ...rooms.slice(1).map(() => '')]);
    data.push(['Total Shortfall (W)', Math.max(0, totalRequired - totalExisting).toFixed(0), ...rooms.slice(1).map(() => '')]);
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    
    // Set column widths - first column wider for labels
    const colWidths = [{wch: 30}];
    rooms.forEach(() => colWidths.push({wch: 20}));
    ws['!cols'] = colWidths;
    
    XLSX.utils.book_append_sheet(wb, ws, 'Heat Loss Assessment');
    
    // Notes sheet (keep this as reference)
    const notes = [
      ['HEAT LOSS CALCULATION NOTES AND GUIDANCE'],
      [],
      ['Design Temperature Explanation:'],
      ['The design temperature is the external ambient temperature used for heat loss calculations.'],
      ['It represents the coldest expected outdoor temperature that the heating system must compensate for.'],
      [],
      ['Common Design Temperatures:'],
      ['@ -3°C: Standard UK design temperature for most regions (BS EN 12831-1:2017)'],
      ['@ -10°C: Used for colder regions, Scotland, or where additional safety margin is required'],
      [],
      ['Mean Water Temperature (MWT) and Delta T:'],
      ['BS EN 442 specifies radiator testing at 75/65°C flow/return = MWT 70°C, ΔT50'],
      [],
      ['Delta T Correction Formula:'],
      ['Actual Output = Rated Output @ ΔT50 × (Actual ΔT ÷ 50)^1.3'],
      [],
      ['U-Values Used (W/m²K):'],
      ['Walls: No external=0, Solid uninsulated=2.1, Solid insulated=0.35, Cavity uninsulated=1.5, Cavity insulated=0.55, Timber=0.3'],
      ['Windows: Single=4.8-5.7, Double=2.8-3.4, Double Low-E=1.8, Triple=2.1-2.6'],
      ['Doors: Solid wood=3.0, Composite=1.8, Glazed single=5.7, Glazed double=3.4'],
      [],
      ['Window Condition Factors:'],
      ['Good=1.0x, Minor wear=1.1x, Damaged seals=1.25x, Failed/Drafty=1.4x'],
      [],
      ['Floor U-Values: Heated below=0, Unheated=0.25, Insulated suspended=0.25, Uninsulated suspended=0.7'],
      ['Insulated concrete ground=0.25, Uninsulated concrete ground=0.7, Over void insulated=0.25, Over void uninsulated=1.0'],
      [],
      ['Ceiling U-Values: Heated above=0, Unheated=0.16, Insulated loft 270mm+=0.16, Partial 100mm=0.3'],
      ['Uninsulated loft=2.3, Insulated flat roof=0.25, Uninsulated flat=1.5, Room in roof insulated=0.2, uninsulated=2.0'],
      [],
      ['Standards: BS EN 12831-1:2017, BS EN 442, CIBSE Domestic Heating Design Guide']
    ];
    
    const ns = XLSX.utils.aoa_to_sheet(notes);
    ns['!cols'] = [{wch: 120}];
    XLSX.utils.book_append_sheet(wb, ns, 'Notes & Guidance');
    
    const fileName = `HeatLoss_${propertyName.replace(/\s+/g, '_') || 'Property'}_${assessmentDate}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-4 pb-32">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
          <div className="flex items-center gap-3 mb-4">
            <Home className="w-8 h-8 text-indigo-600" />
            <h1 className="text-2xl font-bold text-gray-800">Professional Heat Loss Calculator</h1>
          </div>
          
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">Property Name / Address</label>
              <input
                type="text"
                placeholder="Enter property name or address"
                value={propertyName}
                onChange={(e) => setPropertyName(e.target.value)}
                className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none"
              />
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  <User className="w-4 h-4 inline mr-1" />Engineer Name
                </label>
                <input
                  type="text"
                  placeholder="Enter your name"
                  value={engineerName}
                  onChange={(e) => setEngineerName(e.target.value)}
                  className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  <Calendar className="w-4 h-4 inline mr-1" />Assessment Date
                </label>
                <input
                  type="date"
                  value={assessmentDate}
                  onChange={(e) => setAssessmentDate(e.target.value)}
                  className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none"
                />
              </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Building Age</label>
                <select value={buildingAge} onChange={(e) => setBuildingAge(e.target.value)}
                className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                  {Object.keys(BUILDING_AGE).map(age => <option key={age} value={age}>{age}</option>)}
                </select>
              </div>
            
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">Heating Pattern</label>
                <select value={heatingPattern} onChange={(e) => setHeatingPattern(e.target.value)}
                className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                  {Object.keys(HEATING_PATTERN).map(p => <option key={p} value={p}>{p}</option>)}
                </select>
              </div>
            </div>
          </div>
        </div>
      </div>

        {/* Rooms */}
        {rooms.map((room, idx) => (
          <div key={room.id} className="bg-white rounded-lg shadow-lg p-6 mb-6">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-gray-800">Room {idx + 1}</h2>
              {rooms.length > 1 && (
                <button onClick={() => deleteRoom(room.id)} className="p-2 text-red-600 hover:bg-red-50 rounded-lg">
                  <Trash2 className="w-5 h-5" />
                </button>
              )}
            </div>

            <div className="space-y-4">
              {/* Room Type and Temperature */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">Room Type</label>
                  <select value={room.roomType} onChange={(e) => updateRoom(room.id, 'roomType', e.target.value)}
                    className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                    <option value="">Select room type</option>
                    {ROOM_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Required Temp (°C)</label>
                    <input type="number" value={room.requiredTemp}
                      onChange={(e) => updateRoom(room.id, 'requiredTemp', e.target.value)}
                      className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none" />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Design Temp</label>
                    <select value={room.designTemp} onChange={(e) => updateRoom(room.id, 'designTemp', e.target.value)}
                      className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                      <option value="-3">-3°C (UK Standard)</option>
                      <option value="-10">-10°C (Cold Regions)</option>
                    </select>
                  </div>
                </div>
              </div>

              {/* Room Dimensions */}
              <div className="border-t-2 pt-4">
                <h3 className="font-semibold text-gray-800 mb-3">Room Dimensions</h3>
                <div className="grid grid-cols-3 gap-3">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Height (m)</label>
                    <input type="number" step="0.01" value={room.height}
                      onChange={(e) => updateRoom(room.id, 'height', e.target.value)}
                      className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none" />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Length (m)</label>
                    <input type="number" step="0.01" value={room.length}
                      onChange={(e) => updateRoom(room.id, 'length', e.target.value)}
                      className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none" />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Width (m)</label>
                    <input type="number" step="0.01" value={room.width}
                      onChange={(e) => updateRoom(room.id, 'width', e.target.value)}
                      className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none" />
                  </div>
                </div>
                {room.volume && (
                  <div className="mt-3 bg-indigo-50 p-3 rounded-lg">
                    <p className="text-sm font-medium text-indigo-900">Calculated Volume: {room.volume} m³</p>
                  </div>
                )}
              </div>

              {/* Exposed Walls */}
              <div className="border-t-2 pt-4">
                <h3 className="font-semibold text-gray-800 mb-3">Exposed Walls</h3>
                <div className="mb-4">
                  <label className="block text-sm font-medium text-gray-700 mb-2">External Wall Type</label>
                  <select value={room.wallType} onChange={(e) => updateRoom(room.id, 'wallType', e.target.value)}
                    className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                    <option value="">Select wall type</option>
                    {Object.keys(WALL_TYPES).map(t => <option key={t} value={t}>{t} {WALL_TYPES[t] > 0 ? `(U=${WALL_TYPES[t]} W/m²K)` : ''}</option>)}
                  </select>
                </div>
                
                {room.wallType === 'No External Wall' ? (
                  <div className="bg-gray-50 p-4 rounded-lg border-2 border-gray-200">
                    <p className="text-sm text-gray-600">
                      ℹ️ This room has no external walls - wall heat loss will be zero. 
                      Heat loss will only be calculated from windows, doors, floor and ceiling.
                    </p>
                  </div>
                ) : (
                  <>
                    <p className="text-sm text-gray-600 mb-3">Enter up to 3 external walls</p>
                    {[1, 2, 3].map(n => (
                      <div key={n} className="mb-3">
                        <label className="block text-xs font-medium text-gray-600 mb-1">Wall {n}</label>
                        <div className="grid grid-cols-3 gap-2">
                          <div>
                            <input type="number" step="0.01" placeholder="Height (m)" value={room[`wall${n}H`]}
                              onChange={(e) => updateRoom(room.id, `wall${n}H`, e.target.value)}
                              className="w-full px-2 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm" />
                          </div>
                          <div>
                            <input type="number" step="0.01" placeholder="Width (m)" value={room[`wall${n}W`]}
                              onChange={(e) => updateRoom(room.id, `wall${n}W`, e.target.value)}
                              className="w-full px-2 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm" />
                          </div>
                          <div>
                            <input type="text" placeholder="Area (m²)" value={room[`wall${n}Area`]} readOnly
                              className="w-full px-2 py-2 border-2 border-gray-200 rounded-lg bg-gray-50 text-sm font-medium" />
                          </div>
                        </div>
                      </div>
                    ))}
                    
                    {room.totalWallArea && parseFloat(room.totalWallArea) > 0 && (
                      <div className="mt-3 bg-indigo-50 p-3 rounded-lg">
                        <p className="text-sm font-medium text-indigo-900">Total Exposed Wall Area: {room.totalWallArea} m²</p>
                      </div>
                    )}
                  </>
                )}
              </div>

              {/* Windows */}
              <div className="border-t-2 pt-4">
                <h3 className="font-semibold text-gray-800 mb-3">Windows (Enter up to 3 windows)</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Window Type</label>
                    <select value={room.windowType} onChange={(e) => updateRoom(room.id, 'windowType', e.target.value)}
                      className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                      <option value="">Select window type</option>
                      {Object.keys(WINDOW_TYPES).map(t => <option key={t} value={t}>{t} (U={WINDOW_TYPES[t]} W/m²K)</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Window Condition</label>
                    <select value={room.windowCondition} onChange={(e) => updateRoom(room.id, 'windowCondition', e.target.value)}
                      className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                      <option value="">Select condition</option>
                      {Object.keys(WINDOW_CONDITIONS).map(t => (
                        <option key={t} value={t}>{t}</option>
                      ))}
                    </select>
                    {room.windowCondition && room.windowCondition !== 'Good Condition' && (
                      <p className="text-xs text-orange-600 mt-1">
                        ⚠ Adds {((WINDOW_CONDITIONS[room.windowCondition].factor - 1) * 100).toFixed(0)}% to window heat loss + infiltration
                      </p>
                    )}
                  </div>
                </div>
                
                {[1, 2, 3].map(n => (
                  <div key={n} className="mb-3">
                    <label className="block text-xs font-medium text-gray-600 mb-1">Window {n}</label>
                    <div className="grid grid-cols-3 gap-2">
                      <div>
                        <input type="number" step="0.01" placeholder="Height (m)" value={room[`win${n}H`]}
                          onChange={(e) => updateRoom(room.id, `win${n}H`, e.target.value)}
                          className="w-full px-2 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm" />
                      </div>
                      <div>
                        <input type="number" step="0.01" placeholder="Width (m)" value={room[`win${n}W`]}
                          onChange={(e) => updateRoom(room.id, `win${n}W`, e.target.value)}
                          className="w-full px-2 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm" />
                      </div>
                      <div>
                        <input type="text" placeholder="Area (m²)" value={room[`win${n}Area`]} readOnly
                          className="w-full px-2 py-2 border-2 border-gray-200 rounded-lg bg-gray-50 text-sm font-medium" />
                      </div>
                    </div>
                  </div>
                ))}
                
                {room.totalWinArea && parseFloat(room.totalWinArea) > 0 && (
                  <div className="mt-3 bg-indigo-50 p-3 rounded-lg">
                    <p className="text-sm font-medium text-indigo-900">Total Window Area: {room.totalWinArea} m²</p>
                  </div>
                )}
              </div>

              {/* External Doors */}
              <div className="border-t-2 pt-4">
                <h3 className="font-semibold text-gray-800 mb-3">External Doors (Enter up to 2 doors)</h3>
                
                {/* Door 1 */}
                <div className="mb-4">
                  <label className="block text-xs font-medium text-gray-600 mb-2">Door 1 (e.g., Front Door)</label>
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                    <div>
                      <select value={room.door1Type} onChange={(e) => updateRoom(room.id, 'door1Type', e.target.value)}
                        className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm">
                        <option value="">Select door type</option>
                        {Object.keys(DOOR_TYPES).map(t => <option key={t} value={t}>{t} (U={DOOR_TYPES[t]})</option>)}
                      </select>
                    </div>
                    <div>
                      <input type="number" step="0.01" placeholder="Height (m)" value={room.door1H}
                        onChange={(e) => updateRoom(room.id, 'door1H', e.target.value)}
                        className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm" />
                    </div>
                    <div>
                      <input type="number" step="0.01" placeholder="Width (m)" value={room.door1W}
                        onChange={(e) => updateRoom(room.id, 'door1W', e.target.value)}
                        className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm" />
                    </div>
                    <div>
                      <input type="text" placeholder="Area (m²)" value={room.door1Area} readOnly
                        className="w-full px-3 py-2 border-2 border-gray-200 rounded-lg bg-gray-50 text-sm font-medium" />
                    </div>
                  </div>
                </div>
                
                {/* Door 2 */}
                <div className="mb-4">
                  <label className="block text-xs font-medium text-gray-600 mb-2">Door 2 (e.g., Back Door / Patio Door)</label>
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                    <div>
                      <select value={room.door2Type} onChange={(e) => updateRoom(room.id, 'door2Type', e.target.value)}
                        className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm">
                        <option value="">Select door type</option>
                        {Object.keys(DOOR_TYPES).map(t => <option key={t} value={t}>{t} (U={DOOR_TYPES[t]})</option>)}
                      </select>
                    </div>
                    <div>
                      <input type="number" step="0.01" placeholder="Height (m)" value={room.door2H}
                        onChange={(e) => updateRoom(room.id, 'door2H', e.target.value)}
                        className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm" />
                    </div>
                    <div>
                      <input type="number" step="0.01" placeholder="Width (m)" value={room.door2W}
                        onChange={(e) => updateRoom(room.id, 'door2W', e.target.value)}
                        className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none text-sm" />
                    </div>
                    <div>
                      <input type="text" placeholder="Area (m²)" value={room.door2Area} readOnly
                        className="w-full px-3 py-2 border-2 border-gray-200 rounded-lg bg-gray-50 text-sm font-medium" />
                    </div>
                  </div>
                </div>
                
                {room.totalDoorArea && parseFloat(room.totalDoorArea) > 0 && (
                  <div className="mt-3 bg-indigo-50 p-3 rounded-lg">
                    <p className="text-sm font-medium text-indigo-900">Total Door Area: {room.totalDoorArea} m²</p>
                  </div>
                )}
              </div>

              {/* Floor & Ceiling Construction */}
              <div className="border-t-2 pt-4">
                <h3 className="font-semibold text-gray-800 mb-3">Floor & Ceiling Construction</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Floor Type</label>
                    <select value={room.floorType} onChange={(e) => updateRoom(room.id, 'floorType', e.target.value)}
                      className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                      <option value="">Select floor type</option>
                      {Object.keys(FLOOR_TYPES).map(t => (
                        <option key={t} value={t}>
                          {t} {FLOOR_TYPES[t].uValue > 0 ? `(U=${FLOOR_TYPES[t].uValue} W/m²K)` : ''}
                        </option>
                      ))}
                    </select>
                    {room.floorType && FLOOR_TYPES[room.floorType]?.uValue >= 0.7 && (
                      <p className="text-xs text-orange-600 mt-1">
                        ⚠ High heat loss floor - consider insulation
                      </p>
                    )}
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Ceiling / Roof Type</label>
                    <select value={room.ceilingType} onChange={(e) => updateRoom(room.id, 'ceilingType', e.target.value)}
                      className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                      <option value="">Select ceiling type</option>
                      {Object.keys(CEILING_TYPES).map(t => (
                        <option key={t} value={t}>
                          {t} {CEILING_TYPES[t].uValue > 0 ? `(U=${CEILING_TYPES[t].uValue} W/m²K)` : ''}
                        </option>
                      ))}
                    </select>
                    {room.ceilingType && CEILING_TYPES[room.ceilingType]?.uValue >= 1.0 && (
                      <p className="text-xs text-orange-600 mt-1">
                        ⚠ High heat loss ceiling - consider insulation
                      </p>
                    )}
                  </div>
                </div>
              </div>

              {/* Heating System Parameters */}
              <div className="border-t-2 pt-4">
                <div className="bg-blue-50 p-4 rounded-lg">
                  <h3 className="font-semibold text-gray-800 mb-3 flex items-center gap-2">
                    <AlertCircle className="w-5 h-5 text-blue-600" />
                    Heating System Parameters
                  </h3>
                  <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                    <div>
                      <label className="block text-xs font-medium text-gray-700 mb-1">Flow Temp (°C)</label>
                      <input type="number" value={room.flowTemp}
                        onChange={(e) => updateRoom(room.id, 'flowTemp', e.target.value)}
                        className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none" />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-gray-700 mb-1">Return Temp (°C)</label>
                      <input type="number" value={room.returnTemp}
                        onChange={(e) => updateRoom(room.id, 'returnTemp', e.target.value)}
                        className="w-full px-3 py-2 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none" />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-gray-700 mb-1">Mean Water Temp</label>
                      <input type="text" value={room.mwt + '°C'} readOnly
                        className="w-full px-3 py-2 border-2 border-gray-200 rounded-lg bg-gray-100 font-medium" />
                    </div>
                    <div>
                      <label className="block text-xs font-medium text-gray-700 mb-1">Delta T</label>
                      <input type="text" value={'ΔT' + room.deltaT} readOnly
                        className="w-full px-3 py-2 border-2 border-gray-200 rounded-lg bg-yellow-100 font-bold text-center" />
                    </div>
                  </div>
                  <p className="text-xs text-gray-600 mt-2">
                    Standard: Flow 75°C, Return 65°C = MWT 70°C, ΔT50. Adjust if using heat pump or different system.
                  </p>
                </div>
              </div>

              {/* Heat Loss Results */}
              {room.requiredHeat && (
                <div className="border-t-2 pt-4">
                  <div className="bg-green-50 p-4 rounded-lg border-2 border-green-200">
                    <h3 className="font-semibold text-green-800 mb-3">Heat Loss Calculation Results</h3>
                    <div className="grid grid-cols-3 gap-4">
                      <div>
                        <p className="text-xs text-gray-600 mb-1">Fabric Loss</p>
                        <p className="text-lg font-bold text-gray-800">{room.fabricHeatLoss} W</p>
                      </div>
                      <div>
                        <p className="text-xs text-gray-600 mb-1">Ventilation Loss</p>
                        <p className="text-lg font-bold text-gray-800">{room.ventilationHeatLoss} W</p>
                      </div>
                      <div>
                        <p className="text-xs text-gray-600 mb-1">TOTAL Required</p>
                        <p className="text-2xl font-bold text-green-600">{room.requiredHeat} W</p>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Existing Radiator */}
              <div className="border-t-2 pt-4">
                <h3 className="font-semibold text-gray-800 mb-3">Existing Radiator Assessment</h3>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-3">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Radiator Type</label>
                    <select value={room.existRadType} onChange={(e) => updateRoom(room.id, 'existRadType', e.target.value)}
                      className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                      <option value="">Select type</option>
                      {RADIATOR_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Height (mm)</label>
                    <select value={room.existRadH} onChange={(e) => updateRoom(room.id, 'existRadH', e.target.value)}
                      className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                      <option value="">Select height</option>
                      <option value="300">300mm</option>
                      <option value="400">400mm</option>
                      <option value="450">450mm</option>
                      <option value="500">500mm</option>
                      <option value="600">600mm</option>
                      <option value="700">700mm</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">Length (mm)</label>
                    <select value={room.existRadL} onChange={(e) => updateRoom(room.id, 'existRadL', e.target.value)}
                      className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                      <option value="">Select length</option>
                      <option value="400">400mm</option>
                      <option value="500">500mm</option>
                      <option value="600">600mm</option>
                      <option value="700">700mm</option>
                      <option value="800">800mm</option>
                      <option value="900">900mm</option>
                      <option value="1000">1000mm</option>
                      <option value="1100">1100mm</option>
                      <option value="1200">1200mm</option>
                      <option value="1400">1400mm</option>
                      <option value="1600">1600mm</option>
                      <option value="1800">1800mm</option>
                      <option value="2000">2000mm</option>
                    </select>
                  </div>
                </div>
                
                {room.existRadOutputDT50 && (
                  <div className="grid grid-cols-2 gap-4 mt-4">
                    <div className="bg-blue-50 p-4 rounded-lg border-2 border-blue-200">
                      <label className="block text-sm font-medium text-blue-800 mb-2">Output @ ΔT50 (Watts)</label>
                      <p className="text-2xl font-bold text-blue-900">{room.existRadOutputDT50} W</p>
                      <p className="text-xs text-blue-600 mt-1">Standard rated output</p>
                    </div>
                    <div className="bg-purple-50 p-4 rounded-lg border-2 border-purple-200">
                      <label className="block text-sm font-medium text-purple-800 mb-2">Actual Output @ ΔT{room.deltaT}</label>
                      <p className="text-2xl font-bold text-purple-900">{room.existRadActual} W</p>
                      <p className="text-xs text-purple-600 mt-1">Corrected for your system</p>
                    </div>
                  </div>
                )}
                
                {room.existRadActual && room.requiredHeat && (
                  <div className={`mt-4 p-4 rounded-lg border-2 ${
                    parseFloat(room.existRadActual) >= parseFloat(room.requiredHeat) 
                      ? 'bg-green-100 border-green-300' 
                      : 'bg-red-100 border-red-300'
                  }`}>
                    <div className="flex items-center gap-2 mb-2">
                      {parseFloat(room.existRadActual) >= parseFloat(room.requiredHeat) ? (
                        <>
                          <div className="w-6 h-6 bg-green-500 rounded-full flex items-center justify-center text-white font-bold">✓</div>
                          <p className="font-bold text-green-800">Existing Radiator is Sufficient</p>
                        </>
                      ) : (
                        <>
                          <div className="w-6 h-6 bg-red-500 rounded-full flex items-center justify-center text-white font-bold">✗</div>
                          <p className="font-bold text-red-800">Existing Radiator is Undersized</p>
                        </>
                      )}
                    </div>
                    <div className="grid grid-cols-2 gap-4 text-sm">
                      <div>
                        <p className="text-gray-600">Required:</p>
                        <p className="font-bold">{room.requiredHeat} W</p>
                      </div>
                      <div>
                        <p className="text-gray-600">Current Output:</p>
                        <p className="font-bold">{room.existRadActual} W</p>
                      </div>
                    </div>
                    {parseFloat(room.existRadActual) < parseFloat(room.requiredHeat) && (
                      <div className="mt-2 pt-2 border-t border-red-200">
                        <p className="text-gray-600 text-sm">Shortfall:</p>
                        <p className="font-bold text-red-700">
                          {(parseFloat(room.requiredHeat) - parseFloat(room.existRadActual)).toFixed(0)} W deficit
                        </p>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* Engineer Recommendation */}
              <div className="border-t-2 pt-4">
                <h3 className="font-semibold text-gray-800 mb-3">Engineer Recommendation</h3>
                <div className="mb-4">
                  <label className="block text-sm font-medium text-gray-700 mb-2">Action Required</label>
                  <select value={room.recommendation} onChange={(e) => updateRoom(room.id, 'recommendation', e.target.value)}
                    className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                    <option value="Keep as is">Keep as is - No action required</option>
                    <option value="Replace radiator">Replace existing radiator</option>
                    <option value="Add additional radiator">Add additional radiator</option>
                  </select>
                </div>
                
                {room.recommendation !== 'Keep as is' && (
                  <div className="bg-yellow-50 p-4 rounded-lg border-2 border-yellow-200">
                    <h4 className="font-semibold text-gray-800 mb-3">
                      {room.recommendation === 'Replace radiator' ? 'Replacement Radiator Specification' : 'Additional Radiator Specification'}
                    </h4>
                    
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-3">
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">Radiator Type</label>
                        <select value={room.recRadType} onChange={(e) => updateRoom(room.id, 'recRadType', e.target.value)}
                          className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                          <option value="">Select type</option>
                          {RADIATOR_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                        </select>
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">Height (mm)</label>
                        <select value={room.recRadH} onChange={(e) => updateRoom(room.id, 'recRadH', e.target.value)}
                          className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                          <option value="">Select height</option>
                          <option value="300">300mm</option>
                          <option value="400">400mm</option>
                          <option value="450">450mm</option>
                          <option value="500">500mm</option>
                          <option value="600">600mm</option>
                          <option value="700">700mm</option>
                        </select>
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">Length (mm)</label>
                        <select value={room.recRadL} onChange={(e) => updateRoom(room.id, 'recRadL', e.target.value)}
                          className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:border-indigo-500 focus:outline-none">
                          <option value="">Select length</option>
                          <option value="400">400mm</option>
                          <option value="500">500mm</option>
                          <option value="600">600mm</option>
                          <option value="700">700mm</option>
                          <option value="800">800mm</option>
                          <option value="900">900mm</option>
                          <option value="1000">1000mm</option>
                          <option value="1100">1100mm</option>
                          <option value="1200">1200mm</option>
                          <option value="1400">1400mm</option>
                          <option value="1600">1600mm</option>
                          <option value="1800">1800mm</option>
                          <option value="2000">2000mm</option>
                        </select>
                      </div>
                    </div>
                    
                    {room.recRadOutputDT50 && (
                      <div className="grid grid-cols-2 gap-4 mt-4">
                        <div className="bg-blue-50 p-4 rounded-lg border-2 border-blue-200">
                          <label className="block text-sm font-medium text-blue-800 mb-2">Output @ ΔT50 (Watts)</label>
                          <p className="text-2xl font-bold text-blue-900">{room.recRadOutputDT50} W</p>
                          <p className="text-xs text-blue-600 mt-1">Standard rated output</p>
                        </div>
                        <div className="bg-green-50 p-4 rounded-lg border-2 border-green-200">
                          <label className="block text-sm font-medium text-green-800 mb-2">Actual Output @ ΔT{room.deltaT}</label>
                          <p className="text-2xl font-bold text-green-900">{room.recRadActual} W</p>
                          <p className="text-xs text-green-600 mt-1">Corrected for your system</p>
                        </div>
                      </div>
                    )}
                    
                    {room.recRadActual && room.requiredHeat && (
                      <div className="mt-3 p-3 bg-white rounded border-2 border-gray-200">
                        <p className="text-sm text-gray-600 mb-1">Coverage Check:</p>
                        <p className={`font-bold ${parseFloat(room.recRadActual) >= parseFloat(room.requiredHeat) ? 'text-green-600' : 'text-orange-600'}`}>
                          {parseFloat(room.recRadActual) >= parseFloat(room.requiredHeat) 
                            ? `✓ Sufficient - ${((parseFloat(room.recRadActual) / parseFloat(room.requiredHeat) - 1) * 100).toFixed(0)}% oversizing margin`
                            : `⚠ Still undersized by ${(parseFloat(room.requiredHeat) - parseFloat(room.recRadActual)).toFixed(0)} W`
                          }
                        </p>
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          </div>
        ))}
        {/* Footer */}
<div className="text-center text-sm text-gray-500 py-4 mb-20">
  <p>
    This app was coded by Claude and Maria. It is available on{' '}
    <a href="https://github.com/MariaBego56/heat-loss-calculator" target="_blank" rel="noopener noreferrer" className="text-indigo-600 hover:underline">GitHub</a>
    {'https://heatlosscalculator.netlify.app/'}and hosted on Netlify. January 2026
  </p>
</div>

        {/* Fixed Bottom Buttons */}
        <div className="fixed bottom-0 left-0 right-0 bg-white border-t-2 border-gray-300 shadow-2xl p-4 z-50">
          <div className="max-w-6xl mx-auto flex gap-3">
            <button onClick={addRoom}
              className="flex-1 bg-indigo-600 text-white py-4 px-6 rounded-lg font-bold text-lg hover:bg-indigo-700 transition flex items-center justify-center gap-2 shadow-lg">
              <Plus className="w-6 h-6" />
              Add Room
            </button>
            <button onClick={generateExcel}
              className="flex-1 bg-green-600 text-white py-4 px-6 rounded-lg font-bold text-lg hover:bg-green-700 transition flex items-center justify-center gap-2 shadow-lg">
              <Download className="w-6 h-6" />
              Export to Excel
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;