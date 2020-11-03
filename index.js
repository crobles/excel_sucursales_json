const readXlsxFile = require('read-excel-file/node');
 
let obj = {Obj_Tiendas: []};
// File path.
readXlsxFile('./src/lugares_de_pago_0710.xlsx').then((rows) => {
  let largo = rows.length;
  for (var i=2; i<largo-2; i++) {
    if (rows[i][8] === "Abierto") {
      let atributos = {Tienda:{}};
      let latlng = JSON.stringify(rows[i][9]).replace(/['"]+/g, '').split(",");
      atributos.Tienda.Geo_Referencia = { Lat: latlng[0], Lng: latlng[1]};
      atributos.Tienda.Direccion = rows[i][4];
      atributos.Tienda.Nombre_Tienda = rows[i][0];
      atributos.Tienda.Pais = rows[i][1];
      atributos.Tienda.Region = rows[i][2];
      atributos.Tienda.Comuna = rows[i][3];
      let xls_services = rows[i][10];
      let arr_services = xls_services.split("·");
      for (var x=0; x<arr_services.length; x++) {
        if(arr_services[x] === ''){
          arr_services.splice(x,1);
        }
      }
      for (var y=0; y<arr_services.length; y++) {
        arr_services[y]= arr_services[y].replace(/[\r\n]/g, '');
      }
      //if (i === 2) {console.log(arr_services.length);console.log(arr_services);}
      atributos.Tienda.Servicios = arr_services; 
      atributos.Tienda.Horarios = {Lunes_a_Viernes: rows[i][5], Sabado: rows[i][6], Domingo_y_Festivos: rows[i][7]};
      atributos.Tienda.Estado = rows[i][8];
      obj.Obj_Tiendas.push(atributos);
    } 
  }
  console.log(JSON.stringify(obj));
})
 
/* Readable Stream.
readXlsxFile(fs.createReadStream('/path/to/file')).then((rows) => {
  ...
})*/