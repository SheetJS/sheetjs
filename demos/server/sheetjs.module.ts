import { Module } from '@nestjs/common';
import { SheetjsController } from './sheetjs.controller';
import { MulterModule } from '@nestjs/platform-express';

@Module({
  controllers: [SheetjsController],
  imports: [
    MulterModule.register({
      dest: './upload',
    }),
  ],
})
export class SheetjsModule {}
