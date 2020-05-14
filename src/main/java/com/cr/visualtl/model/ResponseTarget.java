package com.cr.visualtl.model;

/**
 * @author 唐剑
 * @ClassName ResponseTarget
 * @Description 作为response中的对象, 回传到前台
 * @date 2020-02-14
 */
public class ResponseTarget {

    private boolean success; //是否成功 true代表成功, false代表失败
    private Object data;    //用于向前台传递对象
    private String msg;    //提示信息

    public ResponseTarget(){

    }

    /**
     * @param success 成功与否
     * @Description 建议用于只传递成功状态的时候
     */
    public ResponseTarget(boolean success) {
        this.success = success;
    }

    /**
     * @param success 成功与否
     * @param data    对象数据
     * @Description 建议用于成功时需要传递对象的时候
     */
    public ResponseTarget(boolean success, Object data) {
        this(success);
        this.data = data;
    }

    /**
     * @param success 成功与否
     * @param msg     传递信息
     * @Description 建议用于需要单独传递信息的场景
     */
    public ResponseTarget(boolean success, String msg) {
        this(success);
        this.msg = msg;
    }

    public ResponseTarget(boolean success, Object data, String msg) {
        this.success = success;
        this.data = data;
        this.msg = msg;
    }

    public Object getData() {
        return data;
    }

    public ResponseTarget setData(Object data) {
        this.data = data;
        return this;
    }

    public boolean isSuccess() {
        return success;
    }

    public void setSuccess(boolean success) {
        this.success = success;
    }

    public String getMsg() {
        return msg;
    }

    public ResponseTarget setMsg(String msg) {
        this.msg = msg;
        return this;
    }
}
