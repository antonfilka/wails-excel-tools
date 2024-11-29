export namespace main {
	
	export class SheetResponse {
	    file1Sheets: string[];
	    file2Sheets: string[];
	
	    static createFrom(source: any = {}) {
	        return new SheetResponse(source);
	    }
	
	    constructor(source: any = {}) {
	        if ('string' === typeof source) source = JSON.parse(source);
	        this.file1Sheets = source["file1Sheets"];
	        this.file2Sheets = source["file2Sheets"];
	    }
	}

}

