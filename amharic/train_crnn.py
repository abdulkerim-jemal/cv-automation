import argparse
from .ocr import train_from_csv

if __name__=='__main__':
    p=argparse.ArgumentParser(description='Train the custom Amharic CNN+BiLSTM+CTC OCR model.')
    p.add_argument('--csv',required=True,help='CSV with image,text columns')
    p.add_argument('--image-root',required=True)
    p.add_argument('--output',default='amharic_crnn.pt')
    p.add_argument('--epochs',type=int,default=20)
    p.add_argument('--batch-size',type=int,default=8)
    args=p.parse_args()
    train_from_csv(args.csv,args.image_root,args.output,args.epochs,args.batch_size)
