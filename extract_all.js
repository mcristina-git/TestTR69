const d = db.getSiblingDB('genieacs').devices.findOne({_id: /BC9424/});

if (!d) {
  print(JSON.stringify({}));
} else {
  function extractAll(obj, prefix = '') {
    let res = {};
    for (let k in obj) {
      if (k.startsWith('_')) continue;
      const currentPath = prefix ? prefix + '.' + k : k;
      const item = obj[k];
      
      if (item && typeof item === 'object') {
        if (item._value !== undefined) {
          res[currentPath] = item._value;
        } else if (item._type !== undefined || item._writable !== undefined || item._timestamp !== undefined) {
          res[currentPath] = "-";
        } else {
          Object.assign(res, extractAll(item, currentPath));
        }
      }
    }
    return res;
  }
  print(JSON.stringify(extractAll(d)));
}
