import { TestBed, inject } from '@angular/core/testing';

import { TransService } from './trans.service';

describe('TransService', () => {
  beforeEach(() => {
    TestBed.configureTestingModule({
      providers: [TransService]
    });
  });

  it('should be created', inject([TransService], (service: TransService) => {
    expect(service).toBeTruthy();
  }));
});
