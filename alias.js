

const dat_path = `${__dirname}/dat`;
const src_path = `${__dirname}/src/js`;

const location = file_name => {
	
	let dir = {
		// data files:	
		"dat.txt": `${dat_path}/dat.txt`,
		"out.json":`${dat_path}/out.json`,
		// module files:
		"input_proc.js":`${src_path}/input_proc`
	};

	let path = `unknown file name: ${file_name}.`
	if (dir.hasOwnProperty(file_name)){
		path = dir[file_name];
	}
	return path;
};

module.exports = {
	location
}