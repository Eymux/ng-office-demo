import { Component, OnInit } from '@angular/core';
import { OfficeService } from '../../services/office.service';

@Component({
  selector: 'app-my-component',
  templateUrl: './my-component.component.html',
  styleUrls: ['./my-component.component.css']
})
export class MyComponentComponent implements OnInit {

    constructor(private office: OfficeService) { }

    private text: string = "Hello World!";

    ngOnInit() {
    }

    onClick() {
        this.office.insertText(this.text)
            .then(() => console.log("Finished!"));
    }
}
