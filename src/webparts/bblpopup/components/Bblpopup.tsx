import * as React from 'react';
import styles from './Bblpopup.module.scss';
import { IBblpopupProps } from './IBblpopupProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import {IBblpopupState} from './IBblpopupState';
//import { IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import * as moment from 'moment';
export default class Bblpopup extends React.Component<IBblpopupProps, IBblpopupState> 
{ 
  constructor(props: IBblpopupProps) {
    super(props);
    
    this.state = {closemodalstate:false, ischecked:false};
    this.donotShowAgain = this.donotShowAgain.bind(this);
    this.closePopUp = this.closePopUp.bind(this);
  }

  public render(): React.ReactElement<IBblpopupProps> {
    const {
      description,
      url,
      eventEndDate,
      eventStartDate    
    } = this.props;
   
    const closemodalfn = ():string =>{
    
      if((today.diff(eventStartDate.value,'days')  > 0) && (today.diff(eventEndDate.value,'days') > 0)){
        localStorage.removeItem('closemodalstate');
        return styles.modalclose; 
      }

      if((today.diff(eventStartDate.value,'days')  < 0) && (today.diff(eventEndDate.value,'days') < 0)){
        localStorage.removeItem('closemodalstate');
        return styles.modalclose; 
      }

         if(this.state.closemodalstate)
         {
          return styles.modalclose;         
         }
         else
         {
          return styles.modalopen; 
         }
    };

    const today = moment();
   
    return (
      <div>
        <div id='bblmodal' className= {closemodalfn()}>
         <div className={styles.modalcontent}>
             <span onClick={this.closePopUp} className={styles.close}>&times;</span>  
             <iframe src={url} width='100%' height='100%' allowFullScreen style={{border:'none'}}   />
             <div className={styles.footermodal}>
              <input type='checkbox' onChange={this.donotShowAgain} />
             <span>{description}   start : { moment(eventStartDate.value).format("DD MMMM YYYY")} end : {moment(eventEndDate.value).format("DD MMMM YYYY")}</span>
             </div>
         </div>
        </div>
      </div>   
    );
  }

  
  public componentDidMount(): void { 

   const closestatepop = localStorage.getItem('closemodalstate');

   if((closestatepop === null) || (closestatepop==='false')){
    this.setState(() => {
      return {closemodalstate: false};
    });
   }else {
    this.setState(() => {
      return {closemodalstate: true};
    });
   }
  }
  private donotShowAgain():void {

   // console.log("show");
    //console.log(this.state.ischecked);
        if(this.state.ischecked){
          //console.log("ischeck true");
           //  console.log(this.state.ischecked);
             this.setState(() => {
              return {ischecked: false};
            }); 
        }else{
        //  console.log("ischeck false");
         // console.log(this.state.ischecked);
          this.setState(() => {
            return {ischecked: true};
          }); 
        } 
          
  }

  private  closePopUp():void {

      if(this.state.ischecked){
        localStorage.setItem('closemodalstate','true');  
      }else{
        localStorage.setItem('closemodalstate','false');  
      }
      this.setState(() => {
        return {closemodalstate: true};
      });
  }
}
