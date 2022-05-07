import React from 'react'; // eslint-disable-next-line
import ReactDOM from 'react-dom'; // eslint-disable-next-line

let defaultState = {
    alertStatus: false,
    alertTip: "提示",
    closeAlert: function(){ },
}

class ReminderBox extends React.Component {

    // 框体宽度
    widthFrame = 700; // eslint-disable-next-line
    // 框体高度
    heightFrame = 300; // eslint-disable-next-line
    // 框体X轴位置
    xFrame = 250; // eslint-disable-next-line
    // 框体X轴位置
    yFrame = 20; // eslint-disable-next-line
    // 框体背景颜色
    backgroundColorFrame = '#00FFFF'; // eslint-disable-next-line
    // 界面状态
    state = { ...defaultState }; // eslint-disable-next-line

    constructor(props) {
        super(props);
    }

    // 关闭弹框
    confirm = () => {
        this.setState({
            alertStatus:false
        })
        this.state.closeAlert();
    }

    // 打开弹框
    openFrame = (options) => {
        options = options || {};
        options.alertStatus = true;
        this.setState({
            ...defaultState,
            ...options
        })
    }



    render() {
        // 框体整体样式
        const container = {
            position: 'absolute',
            marginLeft: this.xFrame,
            marginTop: this.yFrame,
            width: (this.widthFrame + 'px'),
            height: (this.heightFrame + 'px'),
            border: '2px solid blue',
            backgroundColor: this.backgroundColorFrame,
            zIndex: 9999
        }
        // 列表头样式
        const title = {
            height: '25px',
            backgroundColor: 'lightblue'
        }
        // 关闭DIV样式
        const closeCss = {
            marginLeft: (this.widthFrame - 36) + 'px',
            marginTop: (-21) + 'px'
        }
        // 关闭A标签样式
        const closeCssA = {
            cursor: 'default'
        }

        return (
            <div>
                <div style={ this.state.alertStatus? { display : 'block' } : { display : 'none' } }>
                    <div style={ container }>
                        <div style={ title }><div>对公营销提醒</div><div style={ closeCss } onClick={ this.confirm }><a style={ closeCssA }>关闭</a></div></div>
                        <div>{ this.state.alertTip }</div>
                    </div>
                </div>
            </div>
        );
    }
}

let div = document.createElement('div');
let props = { };
document.body.appendChild(div);

let ReminderBoxFrame = ReactDOM.render(React.createElement(ReminderBox, props), div);
export default ReminderBoxFrame