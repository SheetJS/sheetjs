import { Controller, Logger, Post, UploadedFile, UseInterceptors } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { readFile, utils } from 'xlsx';

@Controller('sheetjs')
export class SheetjsController {
  private readonly logger = new Logger(SheetjsController.name);

  @Post('upload-xlsx-file')
  @UseInterceptors(FileInterceptor('file'))
  async uploadXlsxFile(@UploadedFile() file: Express.Multer.File) {
    // Open the uploaded XLSX file and perform SheetJS operations
    const workbook = readFile(file.path);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const output = utils.sheet_to_csv(firstSheet);
    this.logger.log(output);
    return output;
  }
}
