import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { CreateTamplateComponent } from './create-tamplate/create-tamplate.component';

const routes: Routes = [
  {path: 'temp', component: CreateTamplateComponent}
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
