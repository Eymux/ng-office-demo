import { Injectable } from '@angular/core';

@Injectable()
export class OfficeService {

    constructor() { }

    async insertText(text: string): Promise<void> {
        Word.run(context => {
            var doc = context.document;
            doc.body.insertText(text, 'Start');

            return context.sync();
        });
    }
}
