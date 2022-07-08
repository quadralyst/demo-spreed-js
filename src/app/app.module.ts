import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { SpreadSheetsModule } from '@grapecity/spread-sheets-angular';
import { FormsModule, ReactiveFormsModule } from '@angular/forms'; 

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { CreateTamplateComponent } from './create-tamplate/create-tamplate.component';

@NgModule({
  declarations: [
    AppComponent,
    CreateTamplateComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    SpreadSheetsModule,
    FormsModule,
    ReactiveFormsModule,
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
