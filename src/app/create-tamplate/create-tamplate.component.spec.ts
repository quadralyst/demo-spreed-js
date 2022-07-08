import { ComponentFixture, TestBed } from '@angular/core/testing';

import { CreateTamplateComponent } from './create-tamplate.component';

describe('CreateTamplateComponent', () => {
  let component: CreateTamplateComponent;
  let fixture: ComponentFixture<CreateTamplateComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ CreateTamplateComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(CreateTamplateComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
